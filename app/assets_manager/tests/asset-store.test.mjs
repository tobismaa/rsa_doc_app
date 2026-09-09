import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

async function loadStore({ existing = false, user = { uid: 'staff-1' }, failure = null } = {}) {
  const writes = [];
  const stub = {
    db: {}, auth: { currentUser: user },
    collection: (_, name) => name,
    doc: (_, collection, id) => `${collection}/${id}`,
    orderBy: (field, direction) => ({ field, direction }),
    query: (collection, order) => ({ collection, order }),
    getDocs: async query => {
      assert.deepEqual(query, { collection: 'assets', order: { field: 'created_at', direction: 'desc' } });
      if (failure) throw failure;
      return { docs: [{ id: 'IT-01', data: () => ({ name: 'Laptop', id: 'untrusted-field' }) }] };
    },
    serverTimestamp: () => 'SERVER_TIMESTAMP',
    runTransaction: async (_, callback) => {
      if (failure) throw failure;
      await callback({ get: async () => ({ exists: () => existing }), set: (ref, data) => writes.push({ ref, data }) });
    },
  };
  const key = `__assetTest${Math.random().toString(36).slice(2)}`;
  globalThis[key] = stub;
  const source = (await readFile(new URL('../asset-store.js', import.meta.url), 'utf8')).replace(/^import .*;\r?\n/gm, '');
  const prefix = `const { db, auth, collection, doc, getDocs, orderBy, query, runTransaction, serverTimestamp } = globalThis['${key}'];\n`;
  const store = await import(`data:text/javascript;base64,${Buffer.from(prefix + source).toString('base64')}`);
  delete globalThis[key];
  return { store, writes };
}
const payload = { name: 'Laptop', tag: 'IT-01', category: 'Computer equipment', custodian: null, location: null, purchase_cost: 900000, book_value: 750000, purchase_date: null, status: 'Active' };

test('create uses tag identity, user attribution and server timestamps', async () => {
  const { store, writes } = await loadStore();
  await store.createAsset(payload);
  assert.equal(writes[0].ref, 'assets/IT-01');
  assert.equal(writes[0].data.created_by, 'staff-1');
  assert.equal(writes[0].data.created_at, 'SERVER_TIMESTAMP');
  assert.equal(writes[0].data.book_value, 750000);
});
test('duplicate tag never overwrites an existing asset', async () => {
  const { store, writes } = await loadStore({ existing: true });
  await assert.rejects(store.createAsset(payload), { code: 'assets/duplicate-tag' });
  assert.equal(writes.length, 0);
});
test('invalid amounts, paths and whitespace names are rejected before writing', async () => {
  const { store, writes } = await loadStore();
  for (const change of [{ purchase_cost: NaN }, { book_value: -1 }, { tag: '../other' }, { name: '   ' }]) {
    await assert.rejects(store.createAsset({ ...payload, ...change }));
  }
  assert.equal(writes.length, 0);
});
test('unauthenticated creation cannot write', async () => {
  const { store, writes } = await loadStore({ user: null });
  await assert.rejects(store.createAsset(payload), { code: 'assets/not-authorized' });
  assert.equal(writes.length, 0);
});
test('list preserves document identity and reads newest records first', async () => {
  const { store } = await loadStore();
  assert.deepEqual(await store.listAssets(), [{ id: 'IT-01', name: 'Laptop' }]);
});
test('permission failures propagate instead of reporting an empty register or saved asset', async () => {
  const failure = Object.assign(new Error('Denied'), { code: 'permission-denied' });
  const { store, writes } = await loadStore({ failure });
  await assert.rejects(store.listAssets(), { code: 'permission-denied' });
  await assert.rejects(store.createAsset(payload), { code: 'permission-denied' });
  assert.equal(writes.length, 0);
});
