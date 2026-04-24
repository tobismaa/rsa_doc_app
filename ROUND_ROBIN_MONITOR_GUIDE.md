# Round-Robin Distribution Monitoring Guide

## Overview
Your RSA Document App now has a complete round-robin distribution system with admin monitoring, reset capabilities, and testing tools.

---

## 1. HOW THE SYSTEM WORKS

### Distribution Logic
- Each time a customer submits documents (uploader), the system assigns them to reviewers in a rotating manner
- Reviewers are ordered alphabetically by email for deterministic distribution
- Each day at midnight, the counter resets and starts fresh with the first reviewer

**Example:**
```
Day 1:
  Submission 1 → Viewer A
  Submission 2 → Viewer B
  Submission 3 → Viewer C
  Submission 4 → Viewer A (cycle repeats)
  
Day 2 (after midnight):
  Submission 1 → Viewer A (reset)
  Submission 2 → Viewer B
  ...
```

---

## 2. ACCESSING THE MONITORING DASHBOARD

1. **Login as Admin**
2. **Go to Navigation Sidebar**
3. **Click "Round-Robin Monitor"** (between "Reports" and "Audit Log")

---

## 3. MONITORING DASHBOARD SECTIONS

### A. Current Distribution State
Shows real-time status:
- **Last Distribution Viewer**: Who received the most recent submission
- **Current Index**: Position in the viewer rotation (e.g., "2 of 5")
- **Last Reset Date**: When the counter was last reset
- **Total Viewers**: How many active viewers are in the system

### B. Distribution Statistics (Today)
A table showing today's distribution metrics:
- **Viewer Name**: Full name of the viewer
- **Email**: Viewer's email address
- **Assigned Today**: How many submissions were assigned to this viewer today
- **Completed**: How many assigned submissions have been reviewed (approved or rejected)
- **Pending**: How many submissions are still waiting for review

### C. Recent Assignment History
Shows the last 20 assignments with:
- **Timestamp**: When the assignment happened
- **Customer**: Customer name from the submission
- **Assigned To**: Which viewer received it
- **Assigned By**: System or admin email

---

## 4. ADMIN CONTROL BUTTONS

### 🔴 Reset Counter (Today)
**What it does:** Resets the distribution counter to start fresh from the first viewer

**When to use:**
- If the system needs to restart distribution
- After changes to the viewer list
- For daily manual reset (though it auto-resets at midnight)

**Action:**
1. Click "Reset Counter (Today)"
2. Confirm the action
3. System will update all statistics

### 🔵 Test Distribution
**What it does:** Validates that the distribution logic is working correctly

**When to use:**
- Before going live
- After system changes
- To verify everything is calculating correctly

**What it shows:**
```
Distribution Test Results:
================================
Total Viewers: 5
Last Index: 2
Next Index: 3
Next Viewer: viewer3@example.com
Current Date: 2026-03-07
Last Reset Date: 2026-03-07
Test Status: ✅ PASSED
```

### 🟣 View Console Logs
**What it does:** Prints detailed diagnostic information to browser console

**How to view:**
1. Click "View Console Logs"
2. Press **F12** on your keyboard to open Developer Tools
3. Go to the "Console" tab
4. You'll see details like:
   - Current counter state
   - All viewers in the system
   - Recent assignment history

---

## 5. VERIFICATION CHECKLIST

### ✅ Basic Functionality Test
- [ ] Open the Round-Robin Monitor tab
- [ ] Verify "Current Distribution State" shows values
- [ ] Check that "Total Viewers" > 0
- [ ] Confirm "Distribution Statistics" table populates

### ✅ Test Button Verification
- [ ] Click "Test Distribution"
- [ ] Confirmation shows "Test Status: ✅ PASSED"
- [ ] Review the numbers make sense

### ✅ Live Assignment Test
- [ ] Have an uploader submit a document
- [ ] Check the viewer's dashboard
- [ ] Verify the submission appears with status "pending"
- [ ] Refresh the monitoring dashboard
- [ ] "Assigned Today" count should increase

### ✅ Reset Functionality
- [ ] Note the current "Last Distribution Viewer"
- [ ] Click "Reset Counter (Today)"
- [ ] Confirm reset
- [ ] Submit a test document
- [ ] First active viewer should receive it

---

## 6. TROUBLESHOOTING

### Problem: "No reviewers found" error
**Solution:**
- Make sure at least one user with role "reviewer" is created in User Management
- Refresh the page

### Problem: "Assigned Today" count stuck at 0
**Solution:**
- Submissions must have "assignedTo" field set
- Check that submitCustomer() completes successfully
- Look in browser console (F12) for errors

### Problem: Assignment History is empty
**Solution:**
- History collection is created automatically on first assignment
- Submit a test document
- Refresh the monitor
- Check console logs for "✅ Assignment tracked" messages

### Problem: Counter won't reset
**Solution:**
- Check that "counters" collection exists in Firestore
- Try clicking "Reset Counter" again
- Check console logs for error messages

---

## 7. FIRESTORE COLLECTIONS USED

### `counters/roundRobin`
```json
{
  "lastIndex": 2,
  "lastDate": "2026-03-07"
}
```

### `roundRobinAssignments`
Each document records an assignment:
```json
{
  "submissionId": "abc123...",
  "customerName": "John Doe",
  "assignedTo": "viewer2@example.com",
  "assignedBy": "uploader@example.com",
  "assignedAt": "2026-03-07T14:30:00Z",
  "uploadedBy": "uploader@example.com"
}
```

---

## 8. QUICK REFERENCE

| Action | Where | Effect |
|--------|-------|--------|
| Monitor Distribution | Admin > Round-Robin Monitor | View real-time statistics |
| Test System | Click "Test Distribution" | Validate distribution logic |
| Reset Counter | Click "Reset Counter (Today)" | Start fresh distribution |
| View Logs | Click "View Console Logs" + F12 | Detailed system info |
| Check Assignments | Distribution Statistics table | See today's workload |
| View History | Recent Assignment History | Last 20 assignments |

---

## 9. EXPECTED BEHAVIOR

### Normal Daily Flow
```
Morning (after 12:00 AM):
- Counter resets
- First submission → Viewer A
- Second submission → Viewer B
- Third submission → Viewer C
- Fourth submission → Viewer A (cycle)

Afternoon:
- Ratio remains balanced among viewers
- Each viewer gets roughly equal assignments

Evening (before 12:00 AM):
- Final submissions distributed
- Counter state saved
- Awaits next day reset
```

---

## 10. ADMIN RESPONSIBILITIES

- **Daily:** Check monitoring dashboard for balanced distribution
- **Weekly:** Review the statistics to ensure fair load
- **As Needed:** Reset counter if there are changes to viewer list
- **Monthly:** Review assignment history for trends

---

## Questions or Issues?

If you encounter any problems:
1. Check the browser console (F12) for error messages
2. Try the "Test Distribution" button
3. Review the Firestore database in Firebase Console
4. Check that all users have the correct role assignment

---

**Version:** 1.0  
**Last Updated:** March 7, 2026  
**Status:** ✅ Production Ready
