/* global Office */

/**
 * Function file for Office Add-in commands
 * This file handles button click events from the ribbon
 */

Office.onReady(() => {
  // Add-in is ready
  console.log('Office commands ready');
});

/**
 * Quick create pass command
 * Opens the quick pass creation task pane
 */
function quickCreatePass(event) {
  // The task pane is opened automatically by the manifest configuration
  // This function is called when the button is clicked

  // Signal that the function is complete
  event.completed();
}

/**
 * Advanced create pass command
 * Opens the advanced pass creation task pane
 */
function advancedCreatePass(event) {
  // The task pane is opened automatically by the manifest configuration
  // This function is called when the button is clicked

  // Signal that the function is complete
  event.completed();
}

/**
 * Setup command
 * Opens the setup configuration task pane
 */
function openSetup(event) {
  // The task pane is opened automatically by the manifest configuration
  // This function is called when the button is clicked

  // Signal that the function is complete
  event.completed();
}

// Register the functions
if (typeof Office !== 'undefined') {
  Office.actions.associate('quickCreatePass', quickCreatePass);
  Office.actions.associate('advancedCreatePass', advancedCreatePass);
  Office.actions.associate('openSetup', openSetup);
}
