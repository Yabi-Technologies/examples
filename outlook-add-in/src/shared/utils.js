/* global Office */

/**
 * Shared utility functions for PAPI Pass Creator Outlook Add-in
 */

/**
 * Get configuration from Office roaming settings
 * @returns {Promise<Object>} Configuration object
 */
async function getConfig() {
  return new Promise((resolve, reject) => {
    try {
      // Try Office roaming settings first
      if (
        typeof Office !== "undefined" &&
        Office.context &&
        Office.context.roamingSettings
      ) {
        const config = {
          apiUrl: Office.context.roamingSettings.get("apiUrl"),
          clientId: Office.context.roamingSettings.get("clientId"),
          apiToken: Office.context.roamingSettings.get("apiToken"),
          eventId: Office.context.roamingSettings.get("eventId"),
          systemName:
            Office.context.roamingSettings.get("systemName") || "OUTLOOK",
          systemType: Office.context.roamingSettings.get("systemType") || "api",
          sendSmsNewPass:
            Office.context.roamingSettings.get("sendSmsNewPass") !== false,
        };
        resolve(config);
      } else {
        // Fallback to localStorage
        const stored = localStorage.getItem("papiConfig");
        if (stored) {
          resolve(JSON.parse(stored));
        } else {
          resolve({});
        }
      }
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Save configuration to Office roaming settings
 * @param {Object} config - Configuration object
 * @returns {Promise<void>}
 */
async function saveConfig(config) {
  return new Promise((resolve, reject) => {
    try {
      // Try Office roaming settings first
      if (
        typeof Office !== "undefined" &&
        Office.context &&
        Office.context.roamingSettings
      ) {
        Office.context.roamingSettings.set("apiUrl", config.apiUrl);
        Office.context.roamingSettings.set("clientId", config.clientId);
        Office.context.roamingSettings.set("apiToken", config.apiToken);
        Office.context.roamingSettings.set("eventId", config.eventId);
        Office.context.roamingSettings.set("systemName", config.systemName);
        Office.context.roamingSettings.set("systemType", config.systemType);
        Office.context.roamingSettings.set(
          "sendSmsNewPass",
          config.sendSmsNewPass
        );

        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(
              new Error("Failed to save configuration: " + result.error.message)
            );
          }
        });
      } else {
        // Fallback to localStorage
        localStorage.setItem("papiConfig", JSON.stringify(config));
        resolve();
      }
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Create a pass via PAPI API
 * @param {Object} passData - Pass data object
 * @param {Object} config - Configuration object
 * @returns {Promise<Object>} API response
 */
async function createPass(passData, config) {
  try {
    // Prepare the API payload
    const payload = {
      system_name: config.systemName,
      system_type: config.systemType,
      client_id: config.clientId,
      token: config.apiToken,
      send_sms_new_pass: config.sendSmsNewPass,
      passes: [passData],
    };

    // Make the API request
    const response = await fetch(`${config.apiUrl}/passes`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    // Parse the response
    const result = await response.json();

    // Check if the request was successful
    if (response.ok && result.status === "success") {
      return {
        success: true,
        data: result,
        message: "Pass created successfully",
      };
    } else if (result.status === "error") {
      return {
        success: false,
        message: result.message || "Failed to create pass",
        error: result,
      };
    } else if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    return {
      success: false,
      message: "Unexpected response from server",
      error: result,
    };
  } catch (error) {
    console.error("Error creating pass:", error);

    // Provide user-friendly error messages
    let errorMessage = error.message;

    if (
      error.message.includes("Failed to fetch") ||
      error.message.includes("NetworkError")
    ) {
      errorMessage =
        "Network error. Please check your internet connection and API URL.";
    } else if (error.message.includes("401") || error.message.includes("403")) {
      errorMessage =
        "Authentication failed. Please check your API token in Setup.";
    } else if (error.message.includes("404")) {
      errorMessage =
        "API endpoint not found. Please verify the API URL in Setup.";
    } else if (error.message.includes("500")) {
      errorMessage = "Server error. Please contact support or try again later.";
    }

    return {
      success: false,
      message: errorMessage,
      error: error,
    };
  }
}

/**
 * Validate phone number format
 * @param {string} phone - Phone number
 * @returns {boolean} True if valid
 */
function isValidPhone(phone) {
  // Basic validation: 9-15 digits
  return /^[0-9]{9,15}$/.test(phone);
}

/**
 * Validate email format
 * @param {string} email - Email address
 * @returns {boolean} True if valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Format error message for display
 * @param {Error|string} error - Error object or message
 * @returns {string} Formatted error message
 */
function formatError(error) {
  if (typeof error === "string") {
    return error;
  }

  if (error.message) {
    return error.message;
  }

  return "An unknown error occurred";
}

/**
 * Log event for debugging (in production, send to logging service)
 * @param {string} event - Event name
 * @param {Object} data - Event data
 */
function logEvent(event, data) {
  console.log(`[PAPI Pass Creator] ${event}:`, data);

  // In production, you could send this to a logging service
  // Example: sendToLoggingService(event, data);
}

/**
 * Get current timestamp in ISO format
 * @returns {string} ISO timestamp
 */
function getTimestamp() {
  return new Date().toISOString();
}

/**
 * Sanitize string input
 * @param {string} input - Input string
 * @returns {string} Sanitized string
 */
function sanitizeInput(input) {
  if (typeof input !== "string") {
    return input;
  }

  // Remove any potential script tags or HTML
  return input
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#x27;")
    .trim();
}

// Export functions for testing (if in Node.js environment)
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    getConfig,
    saveConfig,
    createPass,
    isValidPhone,
    isValidEmail,
    formatError,
    logEvent,
    getTimestamp,
    sanitizeInput,
  };
}
