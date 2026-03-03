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
 * @property {number}   unitsSold         - Total units sold this day (all channels)
 * @property {number}   soldFromFba       - Units sold from FBA this day
 * @property {number}   soldFromFbm       - Units sold from FBM this day
 * @property {number}   soldFromDtc       - Units sold from DTC this day
 * @property {number}   unfulfilledUnits  - Demand that went unfulfilled (partial OOS)
 * @property {string}   soldFrom          - Channel that fulfilled: FBA|FBM|DTC|OOS
 * @property {number}   conversionRate    - Conversion rate applied this day (%)
 * @property {number}   effectiveVelocity - Primary demand velocity this day
 * @property {number}   effectivePrice    - FBA selling price (used for OOS loss calcs)
 * @property {number}   fbmCostPerUnit    - FBM fulfillment cost per unit (0 if no FBM sales)
 * @property {number}   fbaAvailableEnd   - FBA available at end of day
 * @property {number}   fbaProcessingEnd  - FBA processing at end of day
 * @property {number}   fbmOnHandEnd      - FBM on-hand at end of day
 * @property {number}   dtcEnd            - DTC remaining at end of day
 * @property {string}   channel           - Display channel label (e.g. 'FBA', 'FBA→FBM')
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
  // Subtract customer orders from FBA available — those units are already spoken for.
  // If customer orders exceed available, the remainder eats into the first processing batch.
  var custOrders   = cfg.customerOrders || 0;
  var fbaAvail     = Math.max(0, cfg.fbaAvailable - custOrders);
  var custOverflow = Math.max(0, custOrders - cfg.fbaAvailable); // orders that exceed available
  var fbaProc      = Math.max(0, cfg.fbaProcessing - custOverflow);
  var fbmOnHand    = cfg.fbmOnHand;
  var dtcRemaining = cfg.dtcUnits;

  // ── Pre-compute shipment arrival dates ──
  var spdArrival   = computeArrivalDate(cfg.spdSendDate,   cfg.spdTransitDays);
  var ltlArrival   = computeArrivalDate(cfg.ltlSendDate,   cfg.ltlTransitDays);
  var adhocArrival = computeArrivalDate(cfg.adhocSendDate,  cfg.adhocTransitDays);

  // Inbound Shipped: units already on a truck to Amazon.
  // Arrival = START_DATE + est. days to receive (default 5).
  var inboundShippedUnits = cfg.inboundShippedUnits || 0;
  var inboundShippedArrival = null;
  if (inboundShippedUnits > 0) {
    var shippedTransit = cfg.inboundShippedTransitDays || 5;
    inboundShippedArrival = addDays(START_DATE, shippedTransit);
  }

  // ── Track FBA processing batches (each has an "available on" date) ──
  // A batch: { units: N, availableOn: Date }
  var procBatches = [];

  // If there is starting FBA processing inventory (Inbound: Receiving),
  // it becomes available after the check-in delay from the start date
  if (fbaProc > 0) {
    var initAvailDate = addDays(START_DATE, cfg.fbaCheckinDelay);
    procBatches.push({ units: fbaProc, availableOn: initAvailDate });
  }

  // FC Transfer — units moving between Amazon fulfillment centers.
  // Not sellable until they arrive at the destination FC.
  // Uses the same FBA check-in delay.
  var fcTransfer = cfg.onhandFcTransfer || 0;
  if (fcTransfer > 0) {
    var fcTransferAvailDate = addDays(START_DATE, cfg.fbaCheckinDelay);
    procBatches.push({ units: fcTransfer, availableOn: fcTransferAvailDate });
    fbaProc += fcTransfer;
  }

  // Reserved: FC Processing — units Amazon is verifying, inspecting,
  // or relabeling. This IS the FBA check-in/processing pipeline.
  // Uses the same FBA check-in delay.
  var fcProcessing = cfg.reservedFcProc || 0;
  if (fcProcessing > 0) {
    var fcProcAvailDate = addDays(START_DATE, cfg.fbaCheckinDelay);
    procBatches.push({ units: fcProcessing, availableOn: fcProcAvailDate });
    fbaProc += fcProcessing;
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

    // Inbound Shipped arrival (units Amazon already knows are coming)
    if (inboundShippedArrival && sameDay(today, inboundShippedArrival) && inboundShippedUnits > 0) {
      fbaProc += inboundShippedUnits;
      var shippedAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: inboundShippedUnits, availableOn: shippedAvailOn });
      events.push('Inbound Shipped arrived (' + inboundShippedUnits + ' units) — enters FBA processing');
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

    // ── Step 3: Determine today's FBA velocity ──
    // Check all FBA velocity overrides (first matching override wins)
    var fbaVelocity = cfg.dailyVelocity;
    for (var vo = 0; vo < cfg.velOverrides.length; vo++) {
      var velOv = cfg.velOverrides[vo];
      if (dateInRange(today, velOv.start, velOv.end)) {
        fbaVelocity = velOv.value;
        break;
      }
    }

    // Determine FBM velocity (used when selling from FBM channel)
    var fbmVelocity = cfg.fbmDailyVelocity || fbaVelocity; // fallback to FBA velocity
    if (cfg.fbmVelOverrides) {
      for (var fvo = 0; fvo < cfg.fbmVelOverrides.length; fvo++) {
        var fbmVelOv = cfg.fbmVelOverrides[fvo];
        if (dateInRange(today, fbmVelOv.start, fbmVelOv.end)) {
          fbmVelocity = fbmVelOv.value;
          break;
        }
      }
    }

    // ── Step 4: Fulfillment waterfall — cascading ──
    // Demand cascades through channels: FBA → FBM → DTC → OOS.
    // On transition days (e.g. FBA can't fill full demand), remaining
    // demand spills to the next channel rather than being lost.

    var soldFromFba = 0;
    var soldFromFbm = 0;
    var soldFromDtc = 0;
    var channels    = [];

    // Check if DTC bridge is active as a manual override for this date
    var dtcActive = cfg.dtcStartDate && cfg.dtcEndDate &&
                    dateInRange(today, cfg.dtcStartDate, cfg.dtcEndDate) &&
                    dtcRemaining > 0;

    // Check if FBM start date override forces FBM mode today
    var fbmForced = cfg.fbmStartDateOverride &&
                    today >= cfg.fbmStartDateOverride &&
                    fbmOnHand > 0;

    // Pre-compute effective velocities for both channels
    var fbaEffVel = Math.round(fbaVelocity * (cfg.conversionRate / 100) * 100) / 100;
    var fbmConvRate = (cfg.fbmConversionRate > 0) ? cfg.fbmConversionRate : cfg.conversionRate;

    // FBM demand suppression check: if FBM velocity is more than 75% below FBA,
    // FBM data is likely suppressed by FBA Buy Box dominance (FBA was in stock
    // last year, so FBM sales were ~0). Fall back to FBA × FBM conversion rate.
    var fbmSuppressed = false;
    var fbmEffVel;
    if (fbaVelocity > 0 && fbmVelocity < fbaVelocity * 0.25) {
      fbmSuppressed = true;
      fbmEffVel = Math.round(fbaVelocity * (fbmConvRate / 100) * 100) / 100;
    } else {
      // FBM data is representative — use it directly (no extra conversion rate)
      fbmEffVel = fbmVelocity;
    }

    // Phase 1: FBA — sell from FBA if available (unless FBM forced)
    if (fbaAvail > 0 && !fbmForced) {
      soldFromFba = Math.min(fbaEffVel, fbaAvail);
      fbaAvail = roundInv(fbaAvail - soldFromFba);
      channels.push('FBA');
    }

    // Phase 2: FBM — pick up remaining demand or serve as primary channel
    var fbmDemand = 0;
    if (channels.length > 0 && soldFromFba < fbaEffVel) {
      // FBA tried but couldn't fill all demand — spillover to FBM
      var fbaRemaining = roundInv(fbaEffVel - soldFromFba);
      fbmDemand = Math.min(fbaRemaining, fbmEffVel);
    } else if (channels.length === 0) {
      // No FBA attempt (FBA=0 or fbmForced) — FBM is primary channel
      fbmDemand = fbmEffVel;
    }

    if (fbmDemand > 0 && fbmOnHand > 0) {
      soldFromFbm = Math.min(fbmDemand, fbmOnHand);
      fbmOnHand = roundInv(fbmOnHand - soldFromFbm);
      channels.push('FBM');

      if (soldFromFba > 0) {
        events.push('FBA stock depleted mid-day — ' + fmtNum(soldFromFba) + ' FBA + ' + fmtNum(soldFromFbm) + ' FBM');
      } else if (dayStartFbaAvail > 0 && !fbmForced) {
        events.push('FBA stock depleted — switched to FBM');
      } else if (fbmForced && dayStartFbaAvail > 0) {
        events.push('FBM manual override active — selling from FBM');
      }
    }

    // Phase 3: DTC — pick up anything left after FBA + FBM
    // Primary demand = FBA vel (if FBA was attempted or pure OOS), FBM vel (if FBM was primary)
    var primaryDemand;
    if (channels.length > 0 && channels[0] === 'FBM') {
      primaryDemand = fbmEffVel; // FBM is primary channel (FBA=0 or forced)
    } else {
      primaryDemand = fbaEffVel; // FBA attempted, or pure OOS (use FBA listing demand)
    }
    var remainingAfterFbaFbm = roundInv(primaryDemand - soldFromFba - soldFromFbm);

    if (remainingAfterFbaFbm > 0 && dtcActive) {
      soldFromDtc = Math.min(remainingAfterFbaFbm, dtcRemaining);
      dtcRemaining = roundInv(dtcRemaining - soldFromDtc);
      channels.push('DTC');
      events.push('Fulfilling ' + fmtNum(soldFromDtc) + ' units from DTC bridge');
    }

    // Compute totals
    var unitsSold = soldFromFba + soldFromFbm + soldFromDtc;
    var unfulfilledUnits = roundInv(primaryDemand - unitsSold);

    // Add OOS to channel list if there's unfulfilled demand
    if (unfulfilledUnits > 0 && channels.length > 0) {
      channels.push('OOS');
    }

    // Determine channel label
    var channel;
    if (channels.length === 0) {
      channel = 'OOS';
      events.push('TRUE OOS — no inventory available on any channel');
    } else {
      channel = channels.join('→');
    }

    // Effective velocity = primary channel's demand rate (for reporting/OOS calcs)
    var effectiveVelocity = (channels.length > 0 && channels[0] === 'FBA') ? fbaEffVel : fbmEffVel;
    // If no channels (pure OOS), use FBA velocity to represent lost demand
    if (channels.length === 0) effectiveVelocity = fbaEffVel;

    // Effective price = FBA listing price (used for OOS lost revenue calculations)
    var effectivePrice = cfg.sellingPrice;

    // FBM fulfillment cost (applies when any FBM units were sold)
    var fbmCostPerUnit = (soldFromFbm > 0) ? (cfg.fbmFulfillmentCost || 0) : 0;

    // Check for channel switch back to FBA (if we were on FBM/DTC and FBA became available)
    if ((dayStartFbaAvail === 0) && soldFromFba > 0 && channels[0] === 'FBA') {
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
      soldFromFba:       soldFromFba,
      soldFromFbm:       soldFromFbm,
      soldFromDtc:       soldFromDtc,
      unfulfilledUnits:  unfulfilledUnits,
      soldFrom:          channel,
      conversionRate:    cfg.conversionRate,
      fbmConversionRate: fbmConvRate,
      fbmSuppressed:     fbmSuppressed,
      fbaEffVelocity:    fbaEffVel,
      effectiveVelocity: effectiveVelocity,
      effectivePrice:    effectivePrice,
      fbmCostPerUnit:    fbmCostPerUnit,
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

    // Track the last day of the INITIAL FBA stock run (before first day with no FBA sales).
    // A split day (FBA→FBM) still counts as part of the FBA run since FBA contributed.
    if (!initialFbaEnded && snap.soldFromFba > 0) {
      milestones.lastFbaDay = snap.date;
    } else if (!initialFbaEnded && snap.soldFromFba === 0 && milestones.lastFbaDay !== null) {
      initialFbaEnded = true; // initial FBA run has ended
    }

    // First day on FBM (includes split days where FBM picked up overflow)
    if (!foundFirstFbm && snap.soldFromFbm > 0) {
      milestones.firstFbmDay = snap.date;
      foundFirstFbm = true;
    }

    // First day back on FBA (after being off FBA entirely)
    if (!foundBackOnFba && initialFbaEnded && snap.soldFromFba > 0) {
      milestones.firstBackOnFba = snap.date;
      foundBackOnFba = true;
    }

    // OOS gaps — pure OOS only (no sales from any channel)
    var isPureOos = (snap.channel === 'OOS');
    if (isPureOos) {
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
