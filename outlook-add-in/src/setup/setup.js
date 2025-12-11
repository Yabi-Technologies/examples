/* global Office */

console.log('setup.js loaded');

// Immediate test
alert('Setup script is loading!');

// Try multiple initialization methods
function initializeSetup() {
  console.log('Attempting to initialize setup');
  alert('Initialize function called');
  
  const form = document.getElementById('setupForm');
  const testBtn = document.getElementById('testBtn');
  
  if (!form || !testBtn) {
    console.log('Elements not ready yet, retrying in 100ms');
    setTimeout(initializeSetup, 100);
    return;
  }
  
  alert('Found form and button elements!');
  console.log('Elements found, attaching listeners');
  
  form.addEventListener('submit', handleSave);
  testBtn.addEventListener('click', handleTest);
  
  console.log('Listeners attached successfully');
  alert('Event listeners attached!');

  // Load existing configuration
  loadConfiguration();
}

// Start initialization immediately
initializeSetup();

/**
 * Load existing configuration from storage
 */
async function loadConfiguration() {
  try {
    const config = await getConfig();

    if (config && config.apiUrl) {
      document.getElementById('apiUrl').value = config.apiUrl;
      document.getElementById('clientId').value = config.clientId;
      document.getElementById('apiToken').value = config.apiToken;
      document.getElementById('eventId').value = config.eventId;

      showStatus('Configuration loaded. Update any fields and save.', 'configured');
    }
  } catch (error) {
    console.error('Error loading configuration:', error);
  }
}

/**
 * Handle form submission - save configuration
 */
async function handleSave(event) {
  console.log('handleSave called', event);
  event.preventDefault();
  event.stopPropagation();

  alert('Save button clicked!'); // Debug alert

  const config = {
    apiUrl: document.getElementById('apiUrl').value.trim(),
    clientId: document.getElementById('clientId').value.trim(),
    apiToken: document.getElementById('apiToken').value.trim(),
    eventId: document.getElementById('eventId').value.trim(),
    systemName: 'OUTLOOK',
    systemType: 'api',
    sendSmsNewPass: true
  };

  console.log('Config to save:', config);

  // Validate URL format
  if (!isValidUrl(config.apiUrl)) {
    showMessage('Please enter a valid URL (e.g., https://your-domain.com)', 'error');
    return;
  }

  // Remove trailing slash from URL
  config.apiUrl = config.apiUrl.replace(/\/$/, '');

  try {
    showMessage('Saving configuration...', 'info');

    // Save to Office roaming settings
    await saveConfig(config);

    showMessage('✓ Configuration saved successfully!', 'success');
    showStatus('Add-in is configured and ready to use.', 'configured');

    // Clear form after 2 seconds
    setTimeout(() => {
      hideMessage();
    }, 3000);
  } catch (error) {
    console.error('Error saving configuration:', error);
/**
 * Test the API connection
 */
async function handleTest(event) {
  console.log('handleTest called', event);
  event.preventDefault();
  event.stopPropagation();

  alert('Test button clicked!'); // Debug alert

  const apiUrl = document.getElementById('apiUrl').value.trim();
  const clientId = document.getElementById('clientId').value.trim();
  const apiToken = document.getElementById('apiToken').value.trim();
  const eventId = document.getElementById('eventId').value.trim();

  console.log('Test values:', { apiUrl, clientId, apiToken: '***', eventId });

  if (!apiUrl || !clientId || !apiToken || !eventId) {
    showMessage('Please fill in all fields before testing', 'error');
    return;
  }
  if (!apiUrl || !clientId || !apiToken || !eventId) {
    showMessage('Please fill in all fields before testing', 'error');
    return;
  }

  if (!isValidUrl(apiUrl)) {
    showMessage('Please enter a valid URL', 'error');
    return;
  }

  showMessage('Testing connection...', 'info');
  document.getElementById('testBtn').disabled = true;

  try {
    // Test with a minimal pass creation request (can be test mode)
    const testPayload = {
      system_name: 'OUTLOOK',
      system_type: 'api',
      client_id: clientId,
      token: apiToken,
      send_sms_new_pass: false,
      passes: [
        {
          event_id: eventId,
          mobile: '0000000000', // Test phone
          country_code: '972',
          name: 'TEST',
          last_name: 'CONNECTION',
          category: 11
        }
      ]
    };

    const response = await fetch(`${apiUrl.replace(/\/$/, '')}/passes/test`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(testPayload)
    });

    const result = await response.json();

    if (response.ok || result.status === 'success') {
      showMessage('✓ Connection successful! You can save the configuration.', 'success');
    } else {
      showMessage('Connection test failed: ' + (result.message || 'Unknown error'), 'error');
    }
  } catch (error) {
    console.error('Test failed:', error);
    showMessage('Connection test failed: ' + error.message, 'error');
  } finally {
    document.getElementById('testBtn').disabled = false;
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
}

/**
 * Hide the message
 */
function hideMessage() {
  const messageDiv = document.getElementById('message');
  messageDiv.style.display = 'none';
}

/**
 * Show status message
 */
function showStatus(text, className = '') {
  const statusDiv = document.getElementById('statusMessage');
  statusDiv.textContent = text;
  statusDiv.className = 'status ' + className;
  statusDiv.style.display = 'block';
}

/**
 * Validate URL format
 */
function isValidUrl(string) {
  try {
    const url = new URL(string);
    return url.protocol === 'http:' || url.protocol === 'https:';
  } catch (_) {
    return false;
  }
}
