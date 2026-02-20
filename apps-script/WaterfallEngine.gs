/**
 * WaterfallEngine.gs
 * ---------------------------------------------------------------------------
 * The heart of the forecasting system. For a given SKU's settings, runs
 * a day-by-day sequential simulation from START_DATE to END_DATE.
 *
 * Each day evaluates inventory buckets in priority order:
 *   1. FBA available → sell from here first
 *   2. FBM on-hand  → activates when FBA available hits zero
 *   3. DTC bridge   → manual override within user-specified date range
 *   4. TRUE OOS     → nothing left
 *
 * Inbound shipments (SPD, LTL, ad-hoc) arrive into FBA processing on their
 * calculated arrival date. FBA processing inventory becomes FBA available
 * after the SKU's configurable check-in delay.
 *
 * Returns a structured array of daily snapshots used by both the detail tab
 * renderer and the summary tab builder.
 * ---------------------------------------------------------------------------
 */

/**
 * @typedef {Object} DaySnapshot
 * @property {Date}     date              - Calendar date
 * @property {number}   fbaAvailableStart - FBA available at start of day
 * @property {number}   fbaProcessingStart- FBA processing at start of day
 * @property {number}   fbmOnHandStart    - FBM on-hand at start of day
 * @property {number}   dtcStart          - DTC remaining at start of day
 * @property {number}   unitsSold         - Units sold this day
 * @property {string}   soldFrom          - Channel that fulfilled: FBA|FBM|DTC|OOS
 * @property {number}   conversionRate    - Conversion rate applied this day (%)
 * @property {number}   fbaAvailableEnd   - FBA available at end of day
 * @property {number}   fbaProcessingEnd  - FBA processing at end of day
 * @property {number}   fbmOnHandEnd      - FBM on-hand at end of day
 * @property {number}   dtcEnd            - DTC remaining at end of day
 * @property {string}   channel           - Display channel label
 * @property {string[]} events            - Event log entries for this day
 */

/**
 * Runs the waterfall simulation for one SKU.
 *
 * @param {Object} cfg  Settings object from readSkuSettings()
 * @return {DaySnapshot[]}  Array of daily snapshots, one per forecast day
 */
function runWaterfall(cfg) {
  var dates    = forecastDates();
  var numDays  = dates.length;
  var results  = [];

  // ── Initialize running inventory balances ──
  var fbaAvail     = cfg.fbaAvailable;
  var fbaProc      = cfg.fbaProcessing;
  var fbmOnHand    = cfg.fbmOnHand;
  var dtcRemaining = cfg.dtcUnits;

  // ── Pre-compute shipment arrival dates ──
  var spdArrival   = computeArrivalDate(cfg.spdSendDate,   cfg.spdTransitDays);
  var ltlArrival   = computeArrivalDate(cfg.ltlSendDate,   cfg.ltlTransitDays);
  var adhocArrival = computeArrivalDate(cfg.adhocSendDate,  cfg.adhocTransitDays);

  // ── Track FBA processing batches (each has an "available on" date) ──
  // A batch: { units: N, availableOn: Date }
  var procBatches = [];

  // If there is starting FBA processing inventory, it becomes available
  // after the check-in delay from the start date
  if (fbaProc > 0) {
    var initAvailDate = addDays(START_DATE, cfg.fbaCheckinDelay);
    procBatches.push({ units: fbaProc, availableOn: initAvailDate });
  }

  // ── Day-by-day simulation ──
  for (var d = 0; d < numDays; d++) {
    var today  = dates[d];
    var events = [];

    // Snapshot starting balances (before any events today)
    var dayStartFbaAvail = fbaAvail;
    var dayStartFbaProc  = fbaProc;
    var dayStartFbm      = fbmOnHand;
    var dayStartDtc      = dtcRemaining;

    // ── Step 1: Check if any shipments ARRIVE today ──
    // SPD arrival
    if (spdArrival && sameDay(today, spdArrival) && cfg.spdUnits > 0) {
      fbaProc += cfg.spdUnits;
      var spdAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.spdUnits, availableOn: spdAvailOn });
      events.push('SPD shipment arrived (' + cfg.spdUnits + ' units) — enters FBA processing');
    }

    // LTL arrival
    if (ltlArrival && sameDay(today, ltlArrival) && cfg.ltlUnits > 0) {
      fbaProc += cfg.ltlUnits;
      var ltlAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.ltlUnits, availableOn: ltlAvailOn });
      events.push('LTL shipment arrived (' + cfg.ltlUnits + ' units) — enters FBA processing');
    }

    // Ad-hoc arrival
    if (adhocArrival && sameDay(today, adhocArrival) && cfg.adhocUnits > 0) {
      fbaProc += cfg.adhocUnits;
      var adhocAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.adhocUnits, availableOn: adhocAvailOn });
      events.push('Ad-hoc shipment arrived (' + cfg.adhocUnits + ' units) — enters FBA processing');
    }

    // ── Step 2: Check if any processing batches clear check-in today ──
    for (var b = procBatches.length - 1; b >= 0; b--) {
      if (sameDay(today, procBatches[b].availableOn) || today > procBatches[b].availableOn) {
        // Only process batches that become available exactly today
        // (batches from before START_DATE that should already be available are handled too)
        if (sameDay(today, procBatches[b].availableOn)) {
          var batchUnits = procBatches[b].units;
          fbaAvail += batchUnits;
          fbaProc  -= batchUnits;
          if (fbaProc < 0) fbaProc = 0; // safety clamp
          events.push(batchUnits + ' units cleared FBA processing — now FBA available');
          procBatches.splice(b, 1);
        } else if (today > procBatches[b].availableOn && d === 0) {
          // Edge case: batch was supposed to clear before forecast started
          var batchUnits2 = procBatches[b].units;
          fbaAvail += batchUnits2;
          fbaProc  -= batchUnits2;
          if (fbaProc < 0) fbaProc = 0;
          events.push(batchUnits2 + ' units cleared FBA processing (pre-forecast)');
          procBatches.splice(b, 1);
        }
      }
    }

    // ── Step 3: Determine today's velocity and conversion rate ──
    var velocity = cfg.dailyVelocity;
    if (cfg.velOverrideStart && cfg.velOverrideEnd && cfg.velOverrideValue !== null) {
      if (dateInRange(today, cfg.velOverrideStart, cfg.velOverrideEnd)) {
        velocity = Number(cfg.velOverrideValue) || 0;
      }
    }

    var convRate = cfg.conversionRate;
    if (cfg.crOverrideStart && cfg.crOverrideEnd && cfg.crOverrideValue !== null) {
      if (dateInRange(today, cfg.crOverrideStart, cfg.crOverrideEnd)) {
        convRate = Number(cfg.crOverrideValue) || 0;
      }
    }

    // Apply conversion rate to velocity
    var effectiveVelocity = velocity * (convRate / 100);
    // Round to avoid floating point drift — use banker's rounding to nearest 0.01
    effectiveVelocity = Math.round(effectiveVelocity * 100) / 100;

    // ── Step 4: Fulfillment waterfall — sequential if/else ──
    var unitsSold = 0;
    var channel   = 'OOS';

    // Check if DTC bridge is active as a manual override for this date
    var dtcActive = cfg.dtcStartDate && cfg.dtcEndDate &&
                    dateInRange(today, cfg.dtcStartDate, cfg.dtcEndDate) &&
                    dtcRemaining > 0;

    if (fbaAvail > 0) {
      // Priority 1: FBA available
      channel = 'FBA';
      unitsSold = Math.min(effectiveVelocity, fbaAvail);
      fbaAvail = roundInv(fbaAvail - unitsSold);
    } else if (fbmOnHand > 0) {
      // Priority 2: FBM on-hand (activates the moment FBA hits zero)
      channel = 'FBM';
      unitsSold = Math.min(effectiveVelocity, fbmOnHand);
      fbmOnHand = roundInv(fbmOnHand - unitsSold);
      if (dayStartFbaAvail > 0) {
        events.push('FBA stock depleted — switched to FBM');
      }
    } else if (dtcActive) {
      // Priority 3: DTC bridge (manual override within date range)
      channel = 'DTC';
      unitsSold = Math.min(effectiveVelocity, dtcRemaining);
      dtcRemaining = roundInv(dtcRemaining - unitsSold);
      events.push('Fulfilling from DTC bridge');
    } else {
      // Priority 4: TRUE OOS
      channel = 'OOS';
      unitsSold = 0;
      events.push('TRUE OOS — no inventory available on any channel');
    }

    // Check for channel switch back to FBA (if we were on FBM/DTC and FBA became available)
    if ((dayStartFbaAvail === 0) && (channel === 'FBA')) {
      events.push('FBA stock replenished — switched back to FBA');
    }

    // Log shipment send events (informational)
    if (cfg.spdSendDate && sameDay(today, cfg.spdSendDate) && cfg.spdUnits > 0) {
      events.push('SPD shipment sent (' + cfg.spdUnits + ' units)');
    }
    if (cfg.ltlSendDate && sameDay(today, cfg.ltlSendDate) && cfg.ltlUnits > 0) {
      events.push('LTL shipment sent (' + cfg.ltlUnits + ' units)');
    }
    if (cfg.adhocSendDate && sameDay(today, cfg.adhocSendDate) && cfg.adhocUnits > 0) {
      events.push('Ad-hoc shipment sent (' + cfg.adhocUnits + ' units)');
    }

    // ── Step 5: Record snapshot ──
    results.push({
      date:              today,
      fbaAvailableStart: dayStartFbaAvail,
      fbaProcessingStart:dayStartFbaProc,
      fbmOnHandStart:    dayStartFbm,
      dtcStart:          dayStartDtc,
      unitsSold:         unitsSold,
      soldFrom:          channel,
      conversionRate:    convRate,
      effectiveVelocity: effectiveVelocity,
      fbaAvailableEnd:   fbaAvail,
      fbaProcessingEnd:  fbaProc,
      fbmOnHandEnd:      fbmOnHand,
      dtcEnd:            dtcRemaining,
      channel:           channel,
      events:            events
    });
  }

  return results;
}

/**
 * Computes the arrival date given a send date and transit time in days.
 * Returns null if send date is not set.
 * @param {Date|null} sendDate
 * @param {number}    transitDays
 * @return {Date|null}
 */
function computeArrivalDate(sendDate, transitDays) {
  if (!sendDate) return null;
  return addDays(sendDate, transitDays);
}

/**
 * Rounds an inventory balance to 2 decimal places and clamps to zero.
 * Prevents floating-point drift from accumulating over 71 days of
 * fractional subtraction (e.g., velocity 7 * conversion 33% = 2.31/day).
 * @param {number} val
 * @return {number}
 */
function roundInv(val) {
  var rounded = Math.round(val * 100) / 100;
  return rounded < 0 ? 0 : rounded;
}

/**
 * Adds a number of days to a date and returns a new Date.
 * @param {Date}   d
 * @param {number} days
 * @return {Date}
 */
function addDays(d, days) {
  var result = new Date(d);
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Extracts milestone events from a waterfall result set.
 * Used by the Summary tab to show the milestone table.
 *
 * @param {DaySnapshot[]} snapshots
 * @param {Object}        cfg  - SKU settings for shipment dates
 * @return {Object} milestones
 */
function extractMilestones(snapshots, cfg) {
  var milestones = {
    lastFbaDay:     null,
    firstFbmDay:    null,
    spdArrival:     computeArrivalDate(cfg.spdSendDate, cfg.spdTransitDays),
    ltlArrival:     computeArrivalDate(cfg.ltlSendDate, cfg.ltlTransitDays),
    firstBackOnFba: null,
    oosGaps:        []     // array of {start: Date, end: Date}
  };

  var initialFbaEnded = false;  // true once the initial FBA run ends
  var wasOos          = false;
  var oosStart        = null;
  var foundFirstFbm   = false;
  var foundBackOnFba  = false;

  for (var i = 0; i < snapshots.length; i++) {
    var snap = snapshots[i];

    // Track the last day of the INITIAL FBA stock run (before first switch away)
    if (!initialFbaEnded && snap.channel === 'FBA') {
      milestones.lastFbaDay = snap.date;
    } else if (!initialFbaEnded && snap.channel !== 'FBA' && milestones.lastFbaDay !== null) {
      initialFbaEnded = true; // initial FBA run has ended
    }

    // First day on FBM
    if (!foundFirstFbm && snap.channel === 'FBM') {
      milestones.firstFbmDay = snap.date;
      foundFirstFbm = true;
    }

    // First day back on FBA (after being on FBM/DTC/OOS)
    if (!foundBackOnFba && initialFbaEnded && snap.channel === 'FBA') {
      milestones.firstBackOnFba = snap.date;
      foundBackOnFba = true;
    }

    // OOS gaps
    if (snap.channel === 'OOS') {
      if (!wasOos) {
        oosStart = snap.date;
        wasOos = true;
      }
    } else {
      if (wasOos && oosStart) {
        milestones.oosGaps.push({ start: oosStart, end: snapshots[i - 1].date });
        wasOos = false;
        oosStart = null;
      }
    }
  }

  // Close any open OOS gap
  if (wasOos && oosStart) {
    milestones.oosGaps.push({ start: oosStart, end: snapshots[snapshots.length - 1].date });
  }

  return milestones;
}
