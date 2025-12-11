/* global Office */

let senderInfo = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailInfo();
    document.getElementById('quickPassForm').addEventListener('submit', handleCreatePass);
  }
});

/**
 * Load email sender information
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

      // Update UI
      document.getElementById('senderName').textContent = senderInfo.name || 'Unknown';
      document.getElementById('senderEmail').textContent = senderInfo.email || '';
    } else {
      console.error('Error getting sender info:', result.error);
      document.getElementById('senderName').textContent = 'Unable to load sender info';
    }
  });
}

/**
 * Handle pass creation
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

  const countryCode = document.getElementById('countryCode').value;
  const mobile = document.getElementById('mobile').value.trim();

  // Validate phone number
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

    // Prepare pass data
    const passData = {
      event_id: config.eventId,
      mobile: mobile,
      country_code: countryCode,
      name: senderInfo.firstName,
      last_name: senderInfo.lastName,
      email: senderInfo.email,
      category: 11 // Authorized pass
    };

    // Create pass via API
    const result = await createPass(passData, config);

    if (result.success) {
      showMessage('✓ Pass created successfully!', 'success');

      // Clear form
      document.getElementById('mobile').value = '';

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
