/**
 * getRegisteredActions — the single explicit extension point for the
 * action framework. Returns the ordered array of action descriptors:
 * ICS_CALENDAR_ACTION (src/05-action-ics-import.js),
 * BOOKING_MANAGEMENT_ACTION (src/06-action-booking-com-management.js),
 * TICKETING_PORTALS_ACTION (src/07-action-ticketing-portals.js, v0.6.0
 * quick-260731-tix), TRANSPORT_TICKETS_ACTION
 * (src/08-action-transport-tickets.js, quick-260803-us3),
 * MOJEMENICKA_ACTION (src/09-action-mojemenicka.js, Phase 4 04-01), and
 * MEETINGS_ACTION (src/10-action-meetings.js, quick-260820-g4r).
 *
 * Built inside the function body (not as a top-level const) so cross-file
 * references resolve at call time, after every file's top-level code has
 * run — this avoids GAS file-load-order fragility.
 *
 * To add a future action: create its own file with a descriptor object
 * (name, config, appliesTo(thread), run(thread)), then add it to the
 * array returned here. No other change to the dispatch logic is needed.
 */
function getRegisteredActions() {
  return [ICS_CALENDAR_ACTION, BOOKING_MANAGEMENT_ACTION, TICKETING_PORTALS_ACTION, TRANSPORT_TICKETS_ACTION, MOJEMENICKA_ACTION, MEETINGS_ACTION];
}

/**
 * dispatchActions — runs every registered action against a single thread,
 * in array order. Each action's run is wrapped in its own try/catch so a
 * single misbehaving action can neither abort the loop nor prevent other
 * actions from running nor prevent the thread from being labeled
 * (isolation invariant).
 *
 * ENABLE/DISABLE TOGGLE (cross-cutting, QT-260724-lqi-TOGGLE): before the
 * try block's appliesTo call, any action whose config.enabled is STRICTLY
 * the boolean false is skipped entirely — appliesTo and run are never
 * invoked for it. config.enabled absent, or any other value (including
 * true), is treated as enabled and runs exactly as before. This is the
 * single guard point for every registered action.
 *
 * On a thrown error: logs the failing action name + thread id + error via
 * console.error, then notifies the owner when the action's own config opts
 * in via notifyOnFailure === true.
 *
 * `actions` is an optional injectable array of action descriptors, used by
 * the Node test suite to exercise this function without any GAS global
 * (when supplied, getRegisteredActions() is never called, so no action
 * file's GAS-only globals are referenced). Production callers (processEmails)
 * call dispatchActions(thread) with a single argument, which resolves the
 * working list via getRegisteredActions() exactly as before.
 *
 * Returns { hadError, errors }: hadError is true if any action threw;
 * errors collects { action, error } entries for every action that threw.
 */
function dispatchActions(thread, actions) {
  const actionsToRun = actions || getRegisteredActions();
  const errors = [];

  actionsToRun.forEach(function (action) {
    if (action.config && action.config.enabled === false) {
      return;
    }

    try {
      if (action.appliesTo(thread) === false) {
        return;
      }
      action.run(thread);
    } catch (e) {
      errors.push({ action: action.name, error: String(e) });
      console.error('Action "' + action.name + '" failed on thread ' + thread.getId() + ': ' + e);

      if (action.config && action.config.notifyOnFailure === true) {
        notifyOwnerOfFailure(action.name, thread, e);
      }
    }
  });

  return {
    hadError: errors.length > 0,
    errors: errors,
  };
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { dispatchActions: dispatchActions };
}
