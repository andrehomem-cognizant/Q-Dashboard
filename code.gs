/**
 * @OnlyCurrentDoc
 */

// =================================================================================
// CONFIGURATION
// =================================================================================

// --- Sheet Names ---
const AUDIT_DATA_SHEET = 'Audit, WS 02/06';
const REWORK_DATA_SHEET = 'Rework cases, WS 02/06';

// --- MANAGER ROLES ---
const MANAGER_EMAILS = [
  'inesa.povar@cognizant.com',
  'izabela.goaciniak@cognizant.com',
  'vinesh.nair@cognizant.com',
  'jitesh.amin@cognizant.com',
  'sama.natigzade@cognizant.com',
  'amelia.kalamarska@cognizant.com',
  'karina.dumych@cognizant.com',
  'saktispaul.alexander@cognizant.com',
  'hana.belanova@cognizant.com',
  'yelyzaveta.zatirukha@cognizant.com',
  'neilas.miniotas@cognizant.com',
  'radu.bucurescu@cognizant.com',
  'belen.rodenas@cognizant.com',
  'szymon.gaworski@cognizant.com',
  'serhii.svynarskyi@cognizant.com',
  'elina.mahalova@cognizant.com',
  'stephane.rutkowski@cognizant.com',
  'andre.homem@cognizant.com',
];

// --- Definitive Column Mapping for Audit Sheet ---
const AUDIT_MAP = {
  DATE: 1,
  CASE_ID: 2,
  AGENT_ID: 3,
  SUPPORT_TYPE: 5,
  COUNTRY: 6,
  MENU_REQUEST_TYPE: 10,
  QA_FEEDBACK: 16,
  FIRST_RUBRIC_COL: 17,
  LAST_RUBRIC_COL: 36,
  CRITICAL_ERRORS: 50,
  NON_CRITICAL_ERRORS: 51,
};

// --- Definitive Column Mapping for Rework Sheet ---
const REWORK_MAP = {
  DATE: 1,
  CASE_ID: 2,
  AGENT_ID: 3,
  SUPPORT_TYPE: 5,
  COUNTRY: 6,
  MENU_REQUEST_TYPE: 10,
  VALID_INVALID: 15,
  REWORK_REQUEST: 16,
  FIRST_RUBRIC_COL: 17,
  LAST_RUBRIC_COL: 36, // Corrected to read until column AJ (36)
};


// =================================================================================
// WEB APP & SERVER-SIDE FUNCTIONS
// =================================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Quality Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getFilterOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(AUDIT_DATA_SHEET);
    if (!auditSheet) throw new Error(`Sheet "${AUDIT_DATA_SHEET}" not found.`);

    const markets = getUniqueValues(auditSheet, AUDIT_MAP.COUNTRY);
    const requestTypes = getUniqueValues(auditSheet, AUDIT_MAP.MENU_REQUEST_TYPE);
    const supportTypes = getUniqueValues(auditSheet, AUDIT_MAP.SUPPORT_TYPE);
    return { markets, requestTypes, supportTypes };
  } catch (e) {
    return { error: e.message };
  }
}

function getDashboardData(filters = {}) {
  try {
    Logger.log('-------------------- NEW REQUEST --------------------');
    Logger.log('Starting getDashboardData with filters: ' + JSON.stringify(filters));

    const userEmail = Session.getActiveUser().getEmail();
    const userRole = MANAGER_EMAILS.map(e => e.toLowerCase()).includes(userEmail.toLowerCase()) ? 'manager' : 'agent';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(AUDIT_DATA_SHEET);
    if (!auditSheet) throw new Error(`Sheet "${AUDIT_DATA_SHEET}" not found.`);

    const reworkSheet = ss.getSheetByName(REWORK_DATA_SHEET);
    if (!reworkSheet) throw new Error(`Sheet "${REWORK_DATA_SHEET}" not found.`);

    const auditData = getFilteredData(auditSheet, filters, AUDIT_MAP, 'Audit');
    const reworkData = getFilteredData(reworkSheet, filters, REWORK_MAP, 'Rework');

    const auditRubricHeaders = getRubricHeaders(auditSheet, AUDIT_MAP, 'Audit');
    const reworkRubricHeaders = getRubricHeaders(reworkSheet, REWORK_MAP, 'Rework');
    const allCountries = getUniqueValues(auditSheet, AUDIT_MAP.COUNTRY).filter(c => c !== 'All');

    // Audit data processing
    const auditMetricsByMarket = processAuditMetricsByMarket(auditData);
    const agentMetrics = processAuditMetricsByAgent(auditData, userRole, userEmail);
    const agentTaskDetails = processAgentTaskDetails(auditData, userRole, userEmail);
    const metricsByRequestType = processMetricsByRequestType(auditData);
    const auditCriteria = processCriteria(auditData, auditRubricHeaders, AUDIT_MAP, 'Audit');

    // Rework data processing
    const reworkMetrics = processReworkMetrics(reworkData, allCountries);
    const validReworkData = reworkData.filter(row => row[REWORK_MAP.VALID_INVALID - 1]?.toString().toLowerCase() === 'valid');
    const reworkCriteria = processCriteria(validReworkData, reworkRubricHeaders, REWORK_MAP, 'Rework');
    const reworkAgentMetrics = processReworkMetricsByAgent(reworkData, userRole, userEmail);
    const reworkTaskDetails = processReworkTaskDetails(reworkData, userRole, userEmail);

    const finalReturnObject = {
      auditDataExists: auditData.length > 0,
      reworkDataExists: reworkData.length > 0,
      audits: {
        byMarket: auditMetricsByMarket,
        byAgent: agentMetrics,
        agentTasks: agentTaskDetails,
        criteria: auditCriteria,
        byRequestType: metricsByRequestType
      },
      reworks: {
        metrics: reworkMetrics,
        criteria: reworkCriteria,
        byAgent: reworkAgentMetrics,
        agentTasks: reworkTaskDetails
      }
    };

    Logger.log('-------------------- REQUEST END --------------------');
    return finalReturnObject;

  } catch (e) {
    Logger.log('FATAL ERROR in getDashboardData: ' + e.stack);
    return { error: 'A fatal error occurred on the server: ' + e.message };
  }
}

// =================================================================================
// DATA PROCESSING & AUTOMATED CALCULATIONS
// =================================================================================

function processAuditMetricsByMarket(data) {
  const stats = {};
  data.forEach(row => {
    const market = row[AUDIT_MAP.COUNTRY - 1];
    if (!market) return;
    if (!stats[market]) {
      stats[market] = { casesWithCritical: 0, casesWithNonCritical: 0, caseCount: 0 };
    }
    const criticalErrors = Number(row[AUDIT_MAP.CRITICAL_ERRORS - 1] || 0);
    const nonCriticalErrors = Number(row[AUDIT_MAP.NON_CRITICAL_ERRORS - 1] || 0);
    stats[market].caseCount++;
    if (criticalErrors > 0) stats[market].casesWithCritical++;
    if (nonCriticalErrors > 0) stats[market].casesWithNonCritical++;
  });
  return Object.keys(stats).sort().map(market => {
    const marketStats = stats[market];
    const qualityScore = marketStats.caseCount > 0 ? ((marketStats.caseCount - marketStats.casesWithCritical) / marketStats.caseCount) * 100 : 0;
    return {
      market: market,
      casesAudited: marketStats.caseCount,
      casesWithCritical: marketStats.casesWithCritical,
      casesWithNonCritical: marketStats.casesWithNonCritical,
      qualityScore: qualityScore.toFixed(2) + '%'
    };
  });
}

function processAuditMetricsByAgent(data, userRole, userEmail) {
  let agentData = data;
  if (userRole === 'agent' && userEmail) {
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    agentData = data.filter(row => {
      const sheetEmail = row[AUDIT_MAP.AGENT_ID - 1];
      return sheetEmail && sheetEmail.trim().toLowerCase() === normalizedUserEmail;
    });
  }

  const stats = {};
  agentData.forEach(row => {
    const agentId = row[AUDIT_MAP.AGENT_ID - 1];
    if (!agentId) return;
    if (!stats[agentId]) {
      stats[agentId] = { casesWithCritical: 0, casesWithNonCritical: 0, caseCount: 0 };
    }

    const criticalErrors = Number(row[AUDIT_MAP.CRITICAL_ERRORS - 1] || 0);
    const nonCriticalErrors = Number(row[AUDIT_MAP.NON_CRITICAL_ERRORS - 1] || 0);

    stats[agentId].caseCount++;

    if (criticalErrors > 0) {
      stats[agentId].casesWithCritical++;
    }
    if (nonCriticalErrors > 0) {
      stats[agentId].casesWithNonCritical++;
    }
  });

  return Object.keys(stats).sort().map(agentId => {
    const agentStats = stats[agentId];
    const qualityScore = agentStats.caseCount > 0
      ? ((agentStats.caseCount - agentStats.casesWithCritical) / agentStats.caseCount) * 100
      : 0;

    return {
      agentId: agentId,
      casesAudited: agentStats.caseCount,
      casesWithCritical: agentStats.casesWithCritical,
      casesWithNonCritical: agentStats.casesWithNonCritical,
      qualityScore: qualityScore.toFixed(2) + '%'
    };
  });
}


function processAgentTaskDetails(data, userRole, userEmail) {
  let agentData = data;
  if (userRole === 'agent' && userEmail) {
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    agentData = data.filter(row => {
      const sheetEmail = row[AUDIT_MAP.AGENT_ID - 1];
      return sheetEmail && sheetEmail.trim().toLowerCase() === normalizedUserEmail;
    });
  }
  return agentData.map(row => {
    const criticalErrors = Number(row[AUDIT_MAP.CRITICAL_ERRORS - 1] || 0);
    const nonCriticalErrors = Number(row[AUDIT_MAP.NON_CRITICAL_ERRORS - 1] || 0);
    const quality = (criticalErrors === 0) ? 100 : 0;
    return {
      agentId: row[AUDIT_MAP.AGENT_ID - 1],
      caseId: row[AUDIT_MAP.CASE_ID - 1],
      qaFeedback: row[AUDIT_MAP.QA_FEEDBACK - 1],
      criticalErrors,
      nonCriticalErrors,
      qualityScore: quality.toFixed(0) + '%'
    };
  }).slice(0, 5000);
}

function processReworkMetrics(data, allCountries) {
  const statsByMarket = {};
  allCountries.forEach(country => {
    statsByMarket[country] = {
      validReworks: 0, invalidReworks: 0,
      validWithCritical: 0, validWithNonCritical: 0
    };
  });

  data.forEach(row => {
    const market = row[REWORK_MAP.COUNTRY - 1];
    if (!market || !statsByMarket[market]) return;

    const stats = statsByMarket[market];
    const status = row[REWORK_MAP.VALID_INVALID - 1]?.toString().toLowerCase();

    if (status === 'valid') {
      stats.validReworks++;
      let hasCritical = false;
      let hasNonCritical = false;
      for (let i = REWORK_MAP.FIRST_RUBRIC_COL - 1; i < REWORK_MAP.LAST_RUBRIC_COL; i++) {
        const cellValue = row[i]?.toString().toLowerCase();
        if (cellValue === 'critical') hasCritical = true;
        if (cellValue === 'non-critical') hasNonCritical = true;
      }
      if (hasCritical) stats.validWithCritical++;
      if (hasNonCritical) stats.validWithNonCritical++;
    } else if (status === 'invalid') {
      stats.invalidReworks++;
    }
  });

  return Object.keys(statsByMarket).sort().map(market => {
    const stats = statsByMarket[market];
    return {
      market: market,
      validReworks: stats.validReworks,
      invalidReworks: stats.invalidReworks,
      validWithCritical: stats.validWithCritical,
      validWithNonCritical: stats.validWithNonCritical
    };
  });
}

function processReworkMetricsByAgent(data, userRole, userEmail) {
  let agentData = data;
  if (userRole === 'agent' && userEmail) {
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    agentData = data.filter(row => {
      const sheetEmail = row[REWORK_MAP.AGENT_ID - 1];
      return sheetEmail && sheetEmail.trim().toLowerCase() === normalizedUserEmail;
    });
  }

  const stats = {};
  agentData.forEach(row => {
    const agentId = row[REWORK_MAP.AGENT_ID - 1];
    if (!agentId) return;
    if (!stats[agentId]) {
      stats[agentId] = {
        numberOfReworks: 0,
        valid: 0,
        invalid: 0,
        withCritical: 0,
        withNonCritical: 0
      };
    }

    stats[agentId].numberOfReworks++;
    const status = row[REWORK_MAP.VALID_INVALID - 1]?.toString().toLowerCase();

    if (status === 'valid') {
      stats[agentId].valid++;
      let hasCritical = false;
      let hasNonCritical = false;
      for (let i = REWORK_MAP.FIRST_RUBRIC_COL - 1; i < REWORK_MAP.LAST_RUBRIC_COL; i++) {
        const cellValue = row[i]?.toString().toLowerCase();
        if (cellValue === 'critical') {
          hasCritical = true;
        }
        if (cellValue === 'non-critical') {
          hasNonCritical = true;
        }
      }
      if (hasCritical) {
        stats[agentId].withCritical++;
      }
      if (hasNonCritical) {
        stats[agentId].withNonCritical++;
      }
    } else if (status === 'invalid') {
      stats[agentId].invalid++;
    }
  });

  return Object.keys(stats).sort().map(agentId => {
    const agentStats = stats[agentId];
    return {
      agentId: agentId,
      numberOfReworks: agentStats.numberOfReworks,
      valid: agentStats.valid,
      invalid: agentStats.invalid,
      withCritical: agentStats.withCritical,
      withNonCritical: agentStats.withNonCritical
    };
  });
}

function processReworkTaskDetails(data, userRole, userEmail) {
  let agentData = data;
  if (userRole === 'agent' && userEmail) {
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    agentData = data.filter(row => {
      const sheetEmail = row[REWORK_MAP.AGENT_ID - 1];
      return sheetEmail && sheetEmail.trim().toLowerCase() === normalizedUserEmail;
    });
  }

  return agentData.map(row => {
    let criticalErrors = 0;
    let nonCriticalErrors = 0;
    for (let i = REWORK_MAP.FIRST_RUBRIC_COL - 1; i < REWORK_MAP.LAST_RUBRIC_COL; i++) {
      const cellValue = row[i]?.toString().toLowerCase();
      if (cellValue === 'critical') {
        criticalErrors++;
      } else if (cellValue === 'non-critical') {
        nonCriticalErrors++;
      }
    }

    return {
      agentId: row[REWORK_MAP.AGENT_ID - 1],
      caseId: row[REWORK_MAP.CASE_ID - 1],
      reworkRequest: row[REWORK_MAP.REWORK_REQUEST - 1],
      criticalErrors: criticalErrors,
      nonCriticalErrors: nonCriticalErrors,
    };
  }).slice(0, 5000);
}

function processMetricsByRequestType(data) {
  const stats = { 'Menu Creation': {}, 'Menu Update': {} };
  data.forEach(row => {
    const requestType = row[AUDIT_MAP.MENU_REQUEST_TYPE - 1];
    if (requestType !== 'Menu Creation' && requestType !== 'Menu Update') return;
    const country = row[AUDIT_MAP.COUNTRY - 1];
    if (!country) return;
    if (!stats[requestType][country]) {
      stats[requestType][country] = { casesWithCritical: 0, casesWithNonCritical: 0, caseCount: 0 };
    }
    const countryStats = stats[requestType][country];
    const criticalErrors = Number(row[AUDIT_MAP.CRITICAL_ERRORS - 1] || 0);
    const nonCriticalErrors = Number(row[AUDIT_MAP.NON_CRITICAL_ERRORS - 1] || 0);
    countryStats.caseCount++;
    if (criticalErrors > 0) countryStats.casesWithCritical++;
    if (nonCriticalErrors > 0) countryStats.casesWithNonCritical++;
  });
  const calculateOverall = (typeStats) => {
    const totalCases = Object.values(typeStats).reduce((sum, country) => sum + country.caseCount, 0);
    const totalCritical = Object.values(typeStats).reduce((sum, country) => sum + country.casesWithCritical, 0);
    const qualityScore = totalCases > 0 ? ((totalCases - totalCritical) / totalCases) * 100 : 0;
    return { quality: qualityScore.toFixed(2) + '%' };
  };
  const processResults = (type) => {
    return Object.keys(stats[type]).sort().map(country => {
      const countryStats = stats[type][country];
      const qualityScore = countryStats.caseCount > 0 ? ((countryStats.caseCount - countryStats.casesWithCritical) / countryStats.caseCount) * 100 : 0;
      return {
        country: country,
        casesWithCritical: countryStats.casesWithCritical,
        casesWithNonCritical: countryStats.casesWithNonCritical,
        casesAudited: countryStats.caseCount,
        quality: qualityScore.toFixed(2) + '%'
      };
    });
  };
  return { creation: processResults('Menu Creation'), update: processResults('Menu Update'), creationOverall: calculateOverall(stats['Menu Creation']), updateOverall: calculateOverall(stats['Menu Update']) };
}

function processCriteria(data, rubricHeaders, colMap, sourceName = 'Unknown') {
  const errorCounts = {};
  if (!rubricHeaders) return [];

  const validHeaders = rubricHeaders.map(h => h ? h.trim() : '').filter(h => h !== '');
  Logger.log(`[${sourceName}] Processing criteria against these valid headers: ` + JSON.stringify(validHeaders));

  validHeaders.forEach(header => {
    errorCounts[header] = { nonCritical: 0, critical: 0 };
  });

  const headerIndexMap = {};
  rubricHeaders.forEach((originalHeader, index) => {
    if (originalHeader && originalHeader.trim() !== '') {
      headerIndexMap[originalHeader.trim()] = index;
    }
  });

  data.forEach(row => {
    validHeaders.forEach(header => {
      const colIndex = headerIndexMap[header];
      const cellValue = row[colMap.FIRST_RUBRIC_COL - 1 + colIndex];
      if (cellValue) {
        const valueStr = cellValue.toString().toLowerCase();
        if (valueStr === 'non-critical') {
          errorCounts[header].nonCritical++;
        } else if (valueStr === 'critical') {
          errorCounts[header].critical++;
        }
      }
    });
  });

  Logger.log(`[${sourceName}] Final error counts: ` + JSON.stringify(errorCounts));
  return Object.keys(errorCounts).map(criteria => ({
    criteria: criteria, nonCritical: errorCounts[criteria].nonCritical, critical: errorCounts[criteria].critical
  }));
}

// =================================================================================
// HELPER & UTILITY FUNCTIONS
// =================================================================================

function getFilteredData(sheet, filters, colMap, sourceName = 'Unknown') {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const scriptTimeZone = Session.getScriptTimeZone();

  const startDateStr = filters.startDate ? Utilities.formatDate(new Date(filters.startDate), scriptTimeZone, "yyyyMMdd") : null;
  const endDateStr = filters.endDate ? Utilities.formatDate(new Date(filters.endDate), scriptTimeZone, "yyyyMMdd") : null;

  Logger.log(`[${sourceName}] Filtering data. Total rows: ${allData.length}. Start Date: ${startDateStr || 'None'}. End Date: ${endDateStr || 'None'}`);

  const filteredData = allData.filter(row => {
    // Date filter
    if (startDateStr || endDateStr) {
      const dateValue = row[colMap.DATE - 1];
      if (!dateValue || !(dateValue instanceof Date)) return false;
      const rowDateStr = Utilities.formatDate(dateValue, scriptTimeZone, "yyyyMMdd");
      if (startDateStr && rowDateStr < startDateStr) return false;
      if (endDateStr && rowDateStr > endDateStr) return false;
    }

    // Standard filters
    if (filters.market && filters.market !== 'All' && row[colMap.COUNTRY - 1] !== filters.market) return false;
    if (filters.requestType && filters.requestType !== 'All' && row[colMap.MENU_REQUEST_TYPE - 1] !== filters.requestType) return false;
    if (filters.supportType && filters.supportType !== 'All' && row[colMap.SUPPORT_TYPE - 1] !== filters.supportType) return false;

    // New Error Type Filter
    if (filters.errorType && filters.errorType !== 'All') {
      if (sourceName === 'Audit') {
        if (filters.errorType === 'Critical Errors' && !(Number(row[AUDIT_MAP.CRITICAL_ERRORS - 1] || 0) > 0)) return false;
        if (filters.errorType === 'Non-Critical Errors' && !(Number(row[AUDIT_MAP.NON_CRITICAL_ERRORS - 1] || 0) > 0)) return false;
      } else if (sourceName === 'Rework') {
        let hasError = false;
        const errorStringToFind = filters.errorType === 'Critical Errors' ? 'critical' : 'non-critical';
        for (let i = colMap.FIRST_RUBRIC_COL - 1; i < colMap.LAST_RUBRIC_COL; i++) {
          if (row[i]?.toString().toLowerCase() === errorStringToFind) {
            hasError = true;
            break;
          }
        }
        if (!hasError) return false;
      }
    }

    return true;
  });

  Logger.log(`[${sourceName}] Filtering complete. Rows remaining: ${filteredData.length}`);
  return filteredData.map(row => row.map(cell => cell instanceof Date ? cell.toLocaleDateString() : cell));
}


function getUniqueValues(sheet, columnIndex) {
  if (!sheet || sheet.getLastRow() < 2) return ['All'];
  const data = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getDisplayValues();
  return ['All', ...new Set(data.flat().filter(String).sort())];
}

function getRubricHeaders(sheet, colMap, sourceName = 'Unknown') {
  if (!sheet) return [];
  const numCols = colMap.LAST_RUBRIC_COL - colMap.FIRST_RUBRIC_COL + 1;
  if (numCols < 1) return [];
  const headers = sheet.getRange(1, colMap.FIRST_RUBRIC_COL, 1, numCols).getValues()[0];
  Logger.log(`[${sourceName}] Reading rubric headers from sheet '${sheet.getName()}'. Found: ` + JSON.stringify(headers));
  return headers;
}

