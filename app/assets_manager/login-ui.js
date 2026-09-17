// Keep form navigation and signup available even when the sign-in SDK cannot load.
(() => {
  const loginForm = document.querySelector('#asset-login-form');
  const signupForm = document.querySelector('#asset-signup-form');
  const message = document.querySelector('#login-message');
  const createButton = document.querySelector('#create-account-button');
  const showMessage = (text, success = false) => {
    message.textContent = text;
    message.classList.add('visible');
    message.classList.toggle('success', success);
  };
  const switchForm = signup => {
    loginForm.hidden = signup;
    signupForm.hidden = !signup;
    document.querySelector('#login-title').textContent = signup ? 'Create your account' : 'Welcome back';
    document.querySelector('.intro').textContent = signup
      ? 'Enter your details to create a staff account.'
      : 'Sign in to access your asset workspace.';
    message.classList.remove('visible', 'success');
    document.querySelector('#firebase-fix-link').classList.remove('visible');
    document.querySelector('#login-return-link').classList.remove('visible');
    document.querySelector(signup ? '#display-name' : '#email').focus();
  };
  document.querySelector('#signup-button').addEventListener('click', () => switchForm(true));
  document.querySelector('#back-to-login-button').addEventListener('click', () => switchForm(false));
  document.querySelectorAll('.toggle-password').forEach(toggle => {
    toggle.addEventListener('click', () => {
      const password = toggle.closest('.input-wrap').querySelector('input');
      const show = password.type === 'password';
      password.type = show ? 'text' : 'password';
      toggle.setAttribute('aria-label', show ? 'Hide password' : 'Show password');
      toggle.innerHTML = `<i data-lucide="${show ? 'eye-off' : 'eye'}" aria-hidden="true"></i>`;
      window.lucide?.createIcons();
    });
  });
  signupForm.addEventListener('submit', async event => {
    event.preventDefault();
    if (createButton.disabled || !signupForm.reportValidity()) return;
    const values = new FormData(signupForm);
    const displayName = values.get('displayName').trim();
    if (!displayName) {
      showMessage('Please enter your full name.');
      document.querySelector('#display-name').focus();
      return;
    }
    createButton.disabled = true;
    createButton.querySelector('span').textContent = 'Creating account...';
    message.classList.remove('visible', 'success');
    try {
      const { createStaffAccount } = await import('./signup.js');
      const result = await createStaffAccount({
        displayName,
        email: values.get('email').trim(),
        password: values.get('password'),
      });
      document.querySelector('#email').value = values.get('email').trim();
      signupForm.reset();
      switchForm(false);
      showMessage('Account created successfully. An administrator needs to grant you Asset Manager access.'
        + (result.profileSaved ? '' : ' Your name could not be saved.'), true);
    } catch (error) {
      const errors = {
        'auth/email-already-in-use': 'An account with this email already exists. Please sign in.',
        'auth/invalid-email': 'Please enter a valid email address.',
        'auth/weak-password': 'Please choose a stronger password with at least 6 characters.',
        'auth/password-does-not-meet-requirements': 'Please choose a stronger password with uppercase, lowercase, numbers, and symbols.',
        'auth/operation-not-allowed': 'Email and password signup needs to be enabled in Firebase Authentication.',
        'auth/configuration-not-found': 'Firebase Authentication needs to be enabled for this app.',
        'auth/network-request-failed': 'Check your internet connection and try again.',
        'auth/too-many-requests': 'Too many attempts. Please wait and try again.',
      };
      showMessage(errors[error.code] || (error instanceof TypeError
        ? 'Signup could not load. Check your internet connection and try again.'
        : 'Could not create your account. Please try again.'));
    } finally {
      createButton.disabled = false;
      createButton.querySelector('span').textContent = 'Create account';
    }
  });
  // Block native submission until sign-in handlers have loaded.
  const preventSubmit = event => event.preventDefault();
  loginForm.addEventListener('submit', preventSubmit);
  import('./auth.js').catch(() => {
    document.querySelector('#login-button').disabled = true;
    if (!loginForm.hidden) showMessage('Sign-in could not load. Check your connection and refresh the page.');
  });
})();
