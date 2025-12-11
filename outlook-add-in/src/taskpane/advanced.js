/* global Office */

let senderInfo = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailInfo();
    document.getElementById('advancedPassForm').addEventListener('submit', handleCreatePass);
  }
});

/**
 * Load email sender information and pre-fill form
 */
function loadEmailInfo() {
  Office.context.mailbox.item.from.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      senderInfo = {
        name: result.value.displayName || '',
        email: result.value.emailAddress || ''
      };

      // Split name into first and last name
      const nameParts = senderInfo.name.split(' ');
      senderInfo.firstName = nameParts[0] || '';
      senderInfo.lastName = nameParts.slice(1).join(' ') || '';

      // Pre-fill form
      document.getElementById('firstName').value = senderInfo.firstName;
      document.getElementById('lastName').value = senderInfo.lastName;
      document.getElementById('email').value = senderInfo.email;
    } else {
      console.error('Error getting sender info:', result.error);
    }
  });
}

/**
 * Handle advanced pass creation
 */
async function handleCreatePass(event) {
  event.preventDefault();

  // Check if configured
  const config = await getConfig();
  if (!config || !config.apiUrl) {
    document.getElementById('notConfigured').style.display = 'block';
    showMessage('Please configure the add-in first by clicking the Setup button.', 'error');
    return;
  }

  // Get form values
  const mobile = document.getElementById('mobile').value.trim();
  const countryCode = document.getElementById('countryCode').value;

  // Validate required fields
  if (!mobile || !/^[0-9]{9,15}$/.test(mobile)) {
    showMessage('Please enter a valid phone number (9-15 digits).', 'error');
    return;
  }

  // Disable submit button
  const submitBtn = document.getElementById('submitBtn');
  submitBtn.disabled = true;
  submitBtn.innerHTML = '<span class="loading"></span>Creating pass...';

  try {
    showMessage('Creating pass...', 'info');

    // Prepare pass data with all fields
    const passData = {
      event_id: config.eventId,
      mobile: mobile,
      country_code: countryCode,
      category: 11, // Authorized pass

      // Personal info
      name: document.getElementById('firstName').value.trim() || undefined,
      last_name: document.getElementById('lastName').value.trim() || undefined,
      email: document.getElementById('email').value.trim() || undefined,
      id_number: document.getElementById('idNumber').value.trim() || undefined,

      // Pass details
      code: document.getElementById('passCode').value.trim() || undefined,
      transaction: document.getElementById('transaction').value.trim() || undefined,
      allowed: document.getElementById('allowed').value ? parseFloat(document.getElementById('allowed').value) : undefined,
      single_uses: document.getElementById('singleUses').value ? parseFloat(document.getElementById('singleUses').value) : undefined,
      secured: parseInt(document.getElementById('secured').value),

      // Access control
      areas: document.getElementById('areas').value.trim() || undefined,
      subareas: document.getElementById('subareas').value.trim() || undefined,
      row: document.getElementById('row').value.trim() || undefined,
      seat: document.getElementById('seat').value.trim() || undefined,

      // Options
      transferable: document.getElementById('transferable').checked,
      open_once: document.getElementById('openOnce').checked
    };

    // Remove undefined values
    Object.keys(passData).forEach((key) => {
      if (passData[key] === undefined) {
        delete passData[key];
      }
    });

    // Create pass via API
    const result = await createPass(passData, config);

    if (result.success) {
      showMessage('✓ Pass created successfully!', 'success');

      // Clear optional fields (keep sender info)
      document.getElementById('mobile').value = '';
      document.getElementById('idNumber').value = '';
      document.getElementById('passCode').value = '';
      document.getElementById('transaction').value = '';
      document.getElementById('allowed').value = '';
      document.getElementById('singleUses').value = '';
      document.getElementById('areas').value = '';
      document.getElementById('subareas').value = '';
      document.getElementById('row').value = '';
      document.getElementById('seat').value = '';

      // Reset button after 2 seconds
      setTimeout(() => {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Create Pass';
      }, 2000);
    } else {
      throw new Error(result.message || 'Failed to create pass');
    }
  } catch (error) {
    console.error('Error creating pass:', error);
    showMessage('Error: ' + error.message, 'error');

    submitBtn.disabled = false;
    submitBtn.textContent = 'Create Pass';
  }
}

/**
 * Show a message to the user
 */
function showMessage(text, type) {
  const messageDiv = document.getElementById('message');
  messageDiv.textContent = text;
  messageDiv.className = 'message ' + type;
  messageDiv.style.display = 'block';

  // Auto-hide success messages after 5 seconds
  if (type === 'success') {
    setTimeout(() => {
      messageDiv.style.display = 'none';
    }, 5000);
  }
}
