/**
 * Organization Setup Service
 * Handles initialization and sample data for the Organization module.
 * 
 * The org module now reads personnel from an external Team List spreadsheet
 * and derives structural nodes (task, team, root) from Team Mappings at read time.
 * This setup service manages:
 *   - Team Mappings sheet (local)
 *   - Vacant Positions sheet (local)
 *   - Sample personnel in the external Team List (optional, requires write access)
 */

const OrgSetupService = (function() {
  'use strict';
  
  const MODULE_NAME = 'org';
  // NOTE: CONFIG references must be lazy (inside functions) because
  // config/config.js loads after backend/ files alphabetically.
  function _getTeamMappingsSheet() { return CONFIG.SHEETS.TEAM_MAPPINGS; }
  const VACANT_POSITIONS_SHEET = 'Vacant Positions';
  
  /**
   * Get the bound spreadsheet
   * @private
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  function _getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  /**
   * Ensure a sheet exists with proper headers and formatting
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
   * @param {string} sheetName
   * @param {string[]} headers
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function _ensureSheet(ss, sheetName, headers) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      console.log('[OrgSetupService] Created sheet: ' + sheetName);
    }
    
    // Set headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return sheet;
  }
  
  /**
   * Create a new external Team List spreadsheet, set up headers/formatting,
   * and register it with DataStoreRegistry.
   * @private
   * @param {Object} options
   * @param {string} [options.name] - Spreadsheet name (default: 'Team List - Command Center')
   * @param {string} [options.folderId] - Drive folder ID to move the file into
   * @returns {Object} { success, spreadsheetId, url, sheetName }
   */
  function _createExternalTeamList(options = {}) {
    const spreadsheetName = options.name || 'Team List - Command Center';
    const headers = OrgCrudService.TEAM_LIST_HEADERS;
    
    console.log('[OrgSetupService] Creating external Team List spreadsheet: ' + spreadsheetName);
    
    // 1. Create the spreadsheet
    const newSS = SpreadsheetApp.create(spreadsheetName);
    const spreadsheetId = newSS.getId();
    
    // Move to folder if specified
    if (options.folderId) {
      try {
        const file = DriveApp.getFileById(spreadsheetId);
        const folder = DriveApp.getFolderById(options.folderId);
        folder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
        console.log('[OrgSetupService] Moved spreadsheet to folder: ' + options.folderId);
      } catch (moveErr) {
        console.warn('[OrgSetupService] Could not move to folder: ' + moveErr.message);
      }
    }
    
    // 2. Rename the default "Sheet1" to "Team List"
    const sheetName = CONFIG.SHEETS.TEAM_LIST;
    const defaultSheet = newSS.getSheets()[0];
    defaultSheet.setName(sheetName);
    
    // 3. Write headers
    const sheet = defaultSheet;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 4. Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontWeight('bold')
      .setBackground('#004F87')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setWrap(true);
    sheet.setFrozenRows(1);
    
    // 5. Set column widths for readability
    const colWidths = {
      'Employee Code': 120,
      'Company': 100,
      'Contract': 100,
      'Task': 90,
      'Primary Workstream': 160,
      'Secondary Workstream': 160,
      'First Name': 110,
      'Last Name': 110,
      'Email': 220,
      'Primary Role': 160,
      'Secondary Role': 150,
      'Primary Role Start Date': 140,
      'Contract Personnel Code (CPC)': 120,
      'Heirarchy Identifier (HID)': 140,
      'Unique Personnel ID (UPID)': 150,
      'Supervisor Email': 220,
      'Supervisor UPID': 150,
      'Portfolio Leadership?': 130,
      'Profile Picture': 120,
      'EOD': 100,
      'Personnel Contract Status': 160,
      'Departure Date': 120,
      'Departure Meeting Date': 140,
      'Contract LCAT': 120,
      'Location (City, ST)': 140,
      'Tenure (Days)': 100,
      'Node Type': 100,
      'Active In Org': 100
    };
    headers.forEach((h, i) => {
      if (colWidths[h]) {
        sheet.setColumnWidth(i + 1, colWidths[h]);
      }
    });
    
    // 6. Add data validation for key columns
    const statusCol = headers.indexOf('Personnel Contract Status') + 1;
    if (statusCol > 0) {
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Active', 'Pending EOD', 'Departed'], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, statusCol, 500, 1).setDataValidation(statusRule);
    }
    
    const activeCol = headers.indexOf('Active In Org') + 1;
    if (activeCol > 0) {
      const boolRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, activeCol, 500, 1).setDataValidation(boolRule);
    }
    
    const leadershipCol = headers.indexOf('Portfolio Leadership?') + 1;
    if (leadershipCol > 0) {
      const boolRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, leadershipCol, 500, 1).setDataValidation(boolRule);
    }
    
    // 7. Add conditional formatting for status column
    if (statusCol > 0) {
      const statusRange = sheet.getRange(2, statusCol, 500, 1);
      
      // Active = green
      const activeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Active')
        .setBackground('#D1FAE5')
        .setFontColor('#065F46')
        .setRanges([statusRange])
        .build();
      
      // Departed = gray
      const departedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Departed')
        .setBackground('#E5E7EB')
        .setFontColor('#374151')
        .setRanges([statusRange])
        .build();
      
      // Pending EOD = amber
      const pendingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Pending EOD')
        .setBackground('#FEF3C7')
        .setFontColor('#92400E')
        .setRanges([statusRange])
        .build();
      
      const rules = sheet.getConditionalFormatRules();
      rules.push(activeRule, departedRule, pendingRule);
      sheet.setConditionalFormatRules(rules);
    }
    
    // 8. Add a filter view
    sheet.getRange(1, 1, 1, headers.length).createFilter();
    
    // 9. Register with DataStoreRegistry
    DataStoreRegistry.setSpreadsheetId('teamlist', spreadsheetId);
    console.log('[OrgSetupService] Registered teamlist datastore: ' + spreadsheetId);
    
    return {
      success: true,
      spreadsheetId: spreadsheetId,
      url: newSS.getUrl(),
      sheetName: sheetName
    };
  }
  
  /**
   * Full setup/initialization for the Organization module.
   * Creates the external Team List spreadsheet if not already configured.
   * @param {boolean} includeSampleData - Whether to include sample data
   * @param {Object} [options] - Additional options
   * @param {string} [options.teamListName] - Custom name for external spreadsheet
   * @param {string} [options.folderId] - Drive folder ID to place external spreadsheet
   * @returns {Object} Setup result
   */
  function setup(includeSampleData = false, options = {}) {
    try {
      console.log('[OrgSetupService] Starting setup...');
      
      const ss = _getSpreadsheet();
      
      // 1. Ensure Team Mappings sheet
      const teamHeaders = TeamsCrudService.HEADERS;
      _ensureSheet(ss, _getTeamMappingsSheet(), teamHeaders);
      console.log('[OrgSetupService] Team Mappings sheet ready');
      
      // 2. Ensure Vacant Positions sheet
      const vacantHeaders = TeamsCrudService.VACANT_HEADERS;
      _ensureSheet(ss, VACANT_POSITIONS_SHEET, vacantHeaders);
      console.log('[OrgSetupService] Vacant Positions sheet ready');
      
      // 3. Ensure external Team List spreadsheet exists
      let teamListAccessible = false;
      let teamListCreated = false;
      let teamListUrl = '';
      
      // Check if already configured
      const existingId = DataStoreRegistry.getSpreadsheetId('teamlist');
      if (existingId) {
        // Verify it's actually accessible
        try {
          const tlSheet = DataStoreRegistry.getSheet('teamlist', CONFIG.SHEETS.TEAM_LIST);
          teamListAccessible = !!tlSheet;
          if (teamListAccessible) {
            teamListUrl = DataStoreRegistry.getSpreadsheet('teamlist').getUrl();
          }
          console.log('[OrgSetupService] External Team List accessible: ' + teamListAccessible);
        } catch (tlErr) {
          console.warn('[OrgSetupService] Configured Team List not accessible: ' + tlErr.message);
          console.log('[OrgSetupService] Will recreate external Team List...');
        }
      }
      
      // Create if not configured or not accessible
      if (!teamListAccessible) {
        const createResult = _createExternalTeamList({
          name: options.teamListName || 'Team List - Command Center',
          folderId: options.folderId
        });
        if (createResult.success) {
          teamListAccessible = true;
          teamListCreated = true;
          teamListUrl = createResult.url;
          console.log('[OrgSetupService] Created external Team List: ' + createResult.url);
        } else {
          console.error('[OrgSetupService] Failed to create external Team List');
        }
      }
      
      // 4. Add sample data if requested
      if (includeSampleData) {
        generateSampleData();
      }
      
      // 5. Register with ModuleRegistry
      if (typeof ModuleRegistry !== 'undefined') {
        ModuleRegistry.register({
          id: MODULE_NAME,
          name: 'Org Chart',
          version: '2.0.0',
          healthCheck: healthCheck,
          getSummary: OrgCrudService.getModuleSummary
        });
      }
      
      console.log('[OrgSetupService] Setup complete');
      return {
        success: true,
        message: 'Organization module initialized successfully',
        teamListAccessible: teamListAccessible,
        teamListCreated: teamListCreated,
        teamListUrl: teamListUrl
      };
      
    } catch (error) {
      console.error('[OrgSetupService] Setup failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Health check for the Organization module
   * @returns {Object} Health status
   */
  function healthCheck() {
    const issues = [];
    
    try {
      const ss = _getSpreadsheet();
      
      // Check Team Mappings sheet
      const tmSheet = ss.getSheetByName(_getTeamMappingsSheet());
      if (!tmSheet) {
        issues.push('Team Mappings sheet not found');
      } else {
        const tmData = tmSheet.getDataRange().getValues();
        if (tmData.length < 2) {
          issues.push('Team Mappings sheet has no data (no tasks/teams defined)');
        }
      }
      
      // Check Vacant Positions sheet
      const vpSheet = ss.getSheetByName(VACANT_POSITIONS_SHEET);
      if (!vpSheet) {
        issues.push('Vacant Positions sheet not found');
      }
      
      // Check external Team List accessibility
      try {
        const tlSheet = DataStoreRegistry.getSheet('teamlist', CONFIG.SHEETS.TEAM_LIST);
        if (!tlSheet) {
          issues.push('External Team List sheet not accessible');
        }
      } catch (tlErr) {
        issues.push('Cannot connect to external Team List: ' + tlErr.message);
      }
      
    } catch (error) {
      issues.push('Error during org health check: ' + error.message);
    }
    
    return {
      healthy: issues.length === 0,
      issues: issues,
      module: 'org'
    };
  }
  
  /**
   * Generate sample data for org chart demonstration.
   * Populates:
   *   1. Team Mappings sheet with task/team definitions
   *   2. Vacant Positions sheet with sample vacancies
   *   3. External Team List with sample personnel (if accessible)
   * @returns {Object} Result with record counts
   */
  function generateSampleData() {
    try {
      const ss = _getSpreadsheet();
      const counts = { teamMappings: 0, vacantPositions: 0, personnel: 0 };
      
      // ── 1. Team Mappings sample data ──
      const tmSheet = ss.getSheetByName(_getTeamMappingsSheet()) || _ensureSheet(ss, _getTeamMappingsSheet(), TeamsCrudService.HEADERS);
      const tmLastRow = tmSheet.getLastRow();
      if (tmLastRow > 1) {
        tmSheet.getRange(2, 1, tmLastRow - 1, TeamsCrudService.HEADERS.length).clear();
      }
      
      // [Contract, Task ID, Task Name, Team ID, Team Name, Is Active, Color, Description, Default SLA Threshold, Notify On Escalation, Display Order]
      const teamMappingsData = [
        ['SQuAT', 'TASK-001', 'Task 1 - Program Management',   'TEAM-001', 'Program Management Team',   true, '#9b59b6', 'Core program management',              85, true,  1],
        ['SQuAT', 'TASK-002', 'Task 2 - Acquisition Support',  'TEAM-002', 'APM Team',                  true, '#3498db', 'Acquisition and procurement management', 90, true,  2],
        ['SQuAT', 'TASK-003', 'Task 3 - Portfolio Management',  'TEAM-003', 'PPM Team',                  true, '#1abc9c', 'Portfolio and project management',       85, true,  3],
        ['SQuAT', 'TASK-004', 'Task 4 - Technical Evaluation', 'TEAM-004', 'Technical Evaluation Team', true, '#e74c3c', 'Technical evaluation and QA',            95, true,  4],
        ['SQuAT', 'TASK-005', 'Task 5 - Training Support',     'TEAM-005', 'Training Support Team',     true, '#f39c12', 'Training and documentation',             80, false, 5],
        ['Forward', 'TASK-011', 'Workstream 1 - Strategic Planning', '',    '', true, '#2ecc71', 'Strategic planning workstream',         85, true,  1],
        ['Forward', 'TASK-012', 'Workstream 2 - Implementation',    '',    '', true, '#e67e22', 'Implementation workstream',              90, true,  2],
      ];
      
      if (teamMappingsData.length > 0) {
        tmSheet.getRange(2, 1, teamMappingsData.length, TeamsCrudService.HEADERS.length)
          .setValues(teamMappingsData);
        counts.teamMappings = teamMappingsData.length;
      }
      console.log('[OrgSetupService] Generated ' + counts.teamMappings + ' team mapping records');
      
      // ── 2. Vacant Positions sample data ──
      const vpSheet = ss.getSheetByName(VACANT_POSITIONS_SHEET) || _ensureSheet(ss, VACANT_POSITIONS_SHEET, TeamsCrudService.VACANT_HEADERS);
      const vpLastRow = vpSheet.getLastRow();
      if (vpLastRow > 1) {
        vpSheet.getRange(2, 1, vpLastRow - 1, TeamsCrudService.VACANT_HEADERS.length).clear();
      }
      
      // [Vacant ID, Contract, Task ID, Team ID, Title, Supervisor UPID, Target Hire Date, Requirements, Is Active]
      const vacantData = [
        ['VAC-TASK001-1', 'SQuAT',   'TASK-001', 'TEAM-001', 'Jr. Program Manager',    '310-003',  '2025-04-01', '2+ years PM experience',      true],
        ['VAC-TASK003-1', 'SQuAT',   'TASK-003', 'TEAM-003', 'Developer',               '330-012',  '2025-05-01', '3+ years development experience', true],
        ['VAC-TASK004-1', 'SQuAT',   'TASK-004', 'TEAM-004', 'Team Lead',               '',         '2025-03-15', 'QA Lead with 5+ years experience', true],
      ];
      
      if (vacantData.length > 0) {
        vpSheet.getRange(2, 1, vacantData.length, TeamsCrudService.VACANT_HEADERS.length)
          .setValues(vacantData);
        counts.vacantPositions = vacantData.length;
      }
      console.log('[OrgSetupService] Generated ' + counts.vacantPositions + ' vacant position records');
      
      // ── 3. External Team List sample personnel ──
      try {
        const tlSheet = DataStoreRegistry.getSheet('teamlist', CONFIG.SHEETS.TEAM_LIST);
        if (tlSheet) {
          // Read existing headers from the external sheet
          const headers = OrgCrudService.TEAM_LIST_HEADERS;
          const hIdx = {};
          headers.forEach((h, i) => { hIdx[h] = i; });
          const numCols = headers.length;
          
          // Clear existing data (keep headers)
          const tlLastRow = tlSheet.getLastRow();
          if (tlLastRow > 1) {
            tlSheet.getRange(2, 1, tlLastRow - 1, numCols).clear();
          }
          
          // Build sample personnel rows matching TEAM_LIST_HEADERS order
          const people = _buildSamplePersonnel(hIdx, numCols);
          
          if (people.length > 0) {
            tlSheet.getRange(2, 1, people.length, numCols).setValues(people);
            counts.personnel = people.length;
          }
          console.log('[OrgSetupService] Generated ' + counts.personnel + ' personnel records in Team List');
        } else {
          console.warn('[OrgSetupService] Team List sheet not accessible — skipping personnel sample data');
        }
      } catch (tlErr) {
        console.warn('[OrgSetupService] Could not write sample personnel to external Team List: ' + tlErr.message);
      }
      
      // Invalidate caches after writing sample data
      if (typeof AppCache !== 'undefined') {
        AppCache.invalidate('org:');
        AppCache.invalidate('teams:');
      }
      
      return {
        success: true,
        counts: counts,
        message: `Generated ${counts.teamMappings} team mappings, ${counts.vacantPositions} vacant positions, ${counts.personnel} personnel`
      };
      
    } catch (error) {
      console.error('[OrgSetupService] generateSampleData failed:', error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Build sample personnel rows in TEAM_LIST_HEADERS column order
   * @private
   * @param {Object} hIdx - Header name → column index map
   * @param {number} numCols - Total number of columns
   * @returns {Array<Array>} 2D array of row data
   */
  function _buildSamplePersonnel(hIdx, numCols) {
    const rows = [];
    
    /**
     * Create a row array from a personnel object
     * @param {Object} p - Personnel attributes
     * @returns {Array} Row in TEAM_LIST_HEADERS order
     */
    function makeRow(p) {
      const row = new Array(numCols).fill('');
      row[hIdx['Employee Code']]                  = p.empCode || '';
      row[hIdx['Company']]                        = p.company || 'Hive';
      row[hIdx['Contract']]                       = p.contract || '';
      row[hIdx['Task']]                           = p.task || '';
      row[hIdx['Primary Workstream']]             = p.primaryWorkstream || '';
      row[hIdx['Secondary Workstream']]           = p.secondaryWorkstream || '';
      row[hIdx['First Name']]                     = p.firstName || '';
      row[hIdx['Last Name']]                      = p.lastName || '';
      row[hIdx['Email']]                          = p.email || '';
      row[hIdx['Primary Role']]                   = p.primaryRole || '';
      row[hIdx['Secondary Role']]                 = p.secondaryRole || '';
      row[hIdx['Primary Role Start Date']]        = p.roleStartDate || '';
      row[hIdx['Contract Personnel Code (CPC)']]  = p.cpc || '';
      row[hIdx['Heirarchy Identifier (HID)']]     = p.hid || '';
      row[hIdx['Unique Personnel ID (UPID)']]     = p.upid || '';
      row[hIdx['Supervisor Email']]               = p.supervisorEmail || '';
      row[hIdx['Supervisor UPID']]                = p.supervisorUpid || '';
      row[hIdx['Portfolio Leadership?']]           = p.portfolioLeadership || false;
      row[hIdx['Profile Picture']]                = p.profilePicture || '';
      row[hIdx['EOD']]                            = p.eod || '';
      row[hIdx['Personnel Contract Status']]      = p.status || 'Active';
      row[hIdx['Departure Date']]                 = p.departureDate || '';
      row[hIdx['Departure Meeting Date']]         = p.departureMeetingDate || '';
      row[hIdx['Contract LCAT']]                  = p.lcat || '';
      row[hIdx['Location (City, ST)']]            = p.location || '';
      row[hIdx['Tenure (Days)']]                  = p.tenure || '';
      row[hIdx['Node Type']]                      = p.nodeType || '';
      row[hIdx['Active In Org']]                  = p.activeInOrg !== undefined ? p.activeInOrg : true;
      return row;
    }
    
    // ── SQuAT Executive Leadership (no task — reports to root) ──
    rows.push(makeRow({
      empCode: 'EMP-001', contract: 'SQuAT', firstName: 'John', lastName: 'Richardson',
      email: 'j.richardson@hive.com', primaryRole: 'Executive Leadership',
      cpc: '100', hid: '001', upid: '100-001',
      portfolioLeadership: true, status: 'Active', nodeType: '', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-002', contract: 'SQuAT', firstName: 'Maria', lastName: 'Santos',
      email: 'm.santos@hive.com', primaryRole: 'Executive Leadership', secondaryRole: 'Quality Oversight',
      cpc: '200', hid: '002', upid: '200-002',
      supervisorEmail: 'j.richardson@hive.com', supervisorUpid: '100-001',
      portfolioLeadership: true, status: 'Active', nodeType: '', activeInOrg: true
    }));
    
    // ── Task 1 - Program Management ──
    rows.push(makeRow({
      empCode: 'EMP-003', contract: 'SQuAT', task: 'TASK-001', primaryWorkstream: 'Program Management Team',
      firstName: 'Sarah', lastName: 'Johnson', email: 's.johnson@hive.com',
      primaryRole: 'Program Manager', secondaryRole: 'Scheduling',
      cpc: '310', hid: '003', upid: '310-003',
      supervisorEmail: 'j.richardson@hive.com', supervisorUpid: '100-001',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-004', contract: 'SQuAT', task: 'TASK-001', primaryWorkstream: 'Program Management Team',
      firstName: 'Michael', lastName: 'Chen', email: 'm.chen@hive.com',
      primaryRole: 'Deputy PM',
      cpc: '410', hid: '004', upid: '410-004',
      supervisorEmail: 's.johnson@hive.com', supervisorUpid: '310-003',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-005', contract: 'SQuAT', task: 'TASK-001', primaryWorkstream: 'Program Management Team',
      firstName: 'Alex', lastName: 'Rivera', email: 'a.rivera@hive.com',
      primaryRole: 'Junior PM', secondaryRole: 'Documentation',
      cpc: '410', hid: '005', upid: '410-005',
      supervisorEmail: 'm.chen@hive.com', supervisorUpid: '410-004',
      eod: '2025-02-15', status: 'Pending EOD', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-006', contract: 'SQuAT', task: 'TASK-001', primaryWorkstream: 'Program Management Team',
      firstName: 'Emily', lastName: 'Davis', email: 'e.davis@hive.com',
      primaryRole: 'Program Analyst', secondaryRole: 'Metrics Reporting',
      cpc: '410', hid: '006', upid: '410-006',
      supervisorEmail: 's.johnson@hive.com', supervisorUpid: '310-003',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-007', contract: 'SQuAT', task: 'TASK-001', primaryWorkstream: 'Program Management Team',
      firstName: 'Jason', lastName: 'Park', email: 'j.park@hive.com',
      primaryRole: 'Program Analyst',
      cpc: '410', hid: '007', upid: '410-007',
      supervisorEmail: 'e.davis@hive.com', supervisorUpid: '410-006',
      status: 'Departed', activeInOrg: true
    }));
    
    // ── Task 2 - Acquisition Support ──
    rows.push(makeRow({
      empCode: 'EMP-008', contract: 'SQuAT', task: 'TASK-002', primaryWorkstream: 'APM Team',
      firstName: 'Robert', lastName: 'Wilson', email: 'r.wilson@hive.com',
      primaryRole: 'Acquisition Lead', secondaryRole: 'Contract Writing',
      cpc: '320', hid: '008', upid: '320-008',
      supervisorEmail: 'j.richardson@hive.com', supervisorUpid: '100-001',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-009', contract: 'SQuAT', task: 'TASK-002', primaryWorkstream: 'APM Team',
      firstName: 'Jennifer', lastName: 'Martinez', email: 'j.martinez@hive.com',
      primaryRole: 'Acquisition Specialist',
      cpc: '420', hid: '009', upid: '420-009',
      supervisorEmail: 'r.wilson@hive.com', supervisorUpid: '320-008',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-010', contract: 'SQuAT', task: 'TASK-002', primaryWorkstream: 'APM Team',
      firstName: 'Tyler', lastName: 'Morris', email: 't.morris@hive.com',
      primaryRole: 'Acquisition Specialist',
      cpc: '420', hid: '010', upid: '420-010',
      supervisorEmail: 'j.martinez@hive.com', supervisorUpid: '420-009',
      eod: '2025-03-01', status: 'Pending EOD', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-011', contract: 'SQuAT', task: 'TASK-002', primaryWorkstream: 'APM Team',
      firstName: 'Kevin', lastName: 'Brooks', email: 'k.brooks@hive.com',
      primaryRole: 'Contract Specialist', secondaryRole: 'Legal Review',
      cpc: '420', hid: '011', upid: '420-011',
      supervisorEmail: 'r.wilson@hive.com', supervisorUpid: '320-008',
      status: 'Active', activeInOrg: true
    }));
    
    // ── Task 3 - Portfolio Management ──
    rows.push(makeRow({
      empCode: 'EMP-012', contract: 'SQuAT', task: 'TASK-003', primaryWorkstream: 'PPM Team',
      firstName: 'James', lastName: 'Anderson', email: 'j.anderson@hive.com',
      primaryRole: 'Technical Lead', secondaryRole: 'Architecture',
      cpc: '330', hid: '012', upid: '330-012',
      supervisorEmail: 'j.richardson@hive.com', supervisorUpid: '100-001',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-013', contract: 'SQuAT', task: 'TASK-003', primaryWorkstream: 'PPM Team',
      firstName: 'Patricia', lastName: 'White', email: 'p.white@hive.com',
      primaryRole: 'Developer', secondaryRole: 'Code Review',
      cpc: '430', hid: '013', upid: '430-013',
      supervisorEmail: 'j.anderson@hive.com', supervisorUpid: '330-012',
      status: 'Active', activeInOrg: true
    }));
    
    // ── Task 4 - Technical Evaluation ──
    rows.push(makeRow({
      empCode: 'EMP-014', contract: 'SQuAT', task: 'TASK-004', primaryWorkstream: 'Technical Evaluation Team',
      firstName: 'William', lastName: 'Thompson', email: 'w.thompson@hive.com',
      primaryRole: 'QA Analyst', secondaryRole: 'Test Automation',
      cpc: '440', hid: '014', upid: '440-014',
      supervisorEmail: 'j.richardson@hive.com', supervisorUpid: '100-001',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-015', contract: 'SQuAT', task: 'TASK-004', primaryWorkstream: 'Technical Evaluation Team',
      firstName: 'Rachel', lastName: 'Adams', email: 'r.adams@hive.com',
      primaryRole: 'QA Analyst',
      cpc: '440', hid: '015', upid: '440-015',
      supervisorEmail: 'w.thompson@hive.com', supervisorUpid: '440-014',
      status: 'Active', activeInOrg: true
    }));
    
    // ── Forward Contract ──
    rows.push(makeRow({
      empCode: 'EMP-016', contract: 'Forward', firstName: 'David', lastName: 'Kim',
      email: 'd.kim@hive.com', primaryRole: 'Program Management',
      cpc: '100', hid: '016', upid: '100-016',
      portfolioLeadership: true, status: 'Active', activeInOrg: true
    }));
    
    // Forward Task 1 (TASK-011) - Strategic Planning
    rows.push(makeRow({
      empCode: 'EMP-017', contract: 'Forward', task: 'TASK-011', primaryWorkstream: 'Strategic Planning',
      firstName: 'Lisa', lastName: 'Chen', email: 'l.chen@hive.com',
      primaryRole: 'Strategic Planning', secondaryRole: 'Roadmap Development',
      cpc: '310', hid: '017', upid: '310-017',
      supervisorEmail: 'd.kim@hive.com', supervisorUpid: '100-016',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-018', contract: 'Forward', task: 'TASK-011', primaryWorkstream: 'Strategic Planning',
      firstName: 'Marcus', lastName: 'Johnson', email: 'm.johnson@hive.com',
      primaryRole: 'Planning Analyst',
      cpc: '410', hid: '018', upid: '410-018',
      supervisorEmail: 'l.chen@hive.com', supervisorUpid: '310-017',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-019', contract: 'Forward', task: 'TASK-011', primaryWorkstream: 'Strategic Planning',
      firstName: 'Amanda', lastName: 'Foster', email: 'a.foster@hive.com',
      primaryRole: 'Planning Analyst',
      cpc: '410', hid: '019', upid: '410-019',
      supervisorEmail: 'l.chen@hive.com', supervisorUpid: '310-017',
      eod: '2025-01-20', status: 'Pending EOD', activeInOrg: true
    }));
    
    // Forward Task 2 (TASK-012) - Implementation
    rows.push(makeRow({
      empCode: 'EMP-020', contract: 'Forward', task: 'TASK-012', primaryWorkstream: 'Implementation',
      firstName: 'Brian', lastName: 'Taylor', email: 'b.taylor@hive.com',
      primaryRole: 'Implementation', secondaryRole: 'Change Management',
      cpc: '320', hid: '020', upid: '320-020',
      supervisorEmail: 'd.kim@hive.com', supervisorUpid: '100-016',
      status: 'Active', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-021', contract: 'Forward', task: 'TASK-012', primaryWorkstream: 'Implementation',
      firstName: 'Nicole', lastName: 'Brown', email: 'n.brown@hive.com',
      primaryRole: 'Implementation Specialist',
      cpc: '420', hid: '021', upid: '420-021',
      supervisorEmail: 'b.taylor@hive.com', supervisorUpid: '320-020',
      status: 'Departed', activeInOrg: true
    }));
    rows.push(makeRow({
      empCode: 'EMP-022', contract: 'Forward', task: 'TASK-012', primaryWorkstream: 'Implementation',
      firstName: 'Carlos', lastName: 'Rodriguez', email: 'c.rodriguez@hive.com',
      primaryRole: 'Implementation Specialist', secondaryRole: 'Training',
      cpc: '420', hid: '022', upid: '420-022',
      supervisorEmail: 'b.taylor@hive.com', supervisorUpid: '320-020',
      status: 'Active', activeInOrg: true
    }));
    
    return rows;
  }
  
  /**
   * Standalone method to create/recreate the external Team List spreadsheet.
   * Useful for manual setup or repair without running the full setup flow.
   * @param {Object} [options]
   * @param {string} [options.name] - Custom spreadsheet name
   * @param {string} [options.folderId] - Drive folder ID
   * @param {boolean} [options.includeSampleData] - Also populate sample personnel
   * @returns {Object} Creation result
   */
  function setupExternalTeamList(options = {}) {
    try {
      const result = _createExternalTeamList(options);
      
      if (result.success && options.includeSampleData) {
        // Populate with sample personnel
        try {
          const tlSheet = DataStoreRegistry.getSheet('teamlist', CONFIG.SHEETS.TEAM_LIST);
          if (tlSheet) {
            const headers = OrgCrudService.TEAM_LIST_HEADERS;
            const hIdx = {};
            headers.forEach((h, i) => { hIdx[h] = i; });
            const numCols = headers.length;
            
            const people = _buildSamplePersonnel(hIdx, numCols);
            if (people.length > 0) {
              tlSheet.getRange(2, 1, people.length, numCols).setValues(people);
              result.personnelCount = people.length;
            }
            console.log('[OrgSetupService] Populated Team List with ' + people.length + ' sample personnel');
          }
        } catch (dataErr) {
          console.warn('[OrgSetupService] Created sheet but could not add sample data: ' + dataErr.message);
          result.sampleDataError = dataErr.message;
        }
      }
      
      return result;
    } catch (error) {
      console.error('[OrgSetupService] setupExternalTeamList failed:', error);
      return { success: false, error: error.message };
    }
  }
  
  // Public API
  return {
    setup: setup,
    healthCheck: healthCheck,
    generateSampleData: generateSampleData,
    setupExternalTeamList: setupExternalTeamList
  };
})();

// Auto-register with ModuleRegistry on load
(function() {
  if (typeof ModuleRegistry !== 'undefined') {
    try {
      ModuleRegistry.register({
        id: 'org',
        name: 'Org Chart',
        version: '2.0.0',
        healthCheck: OrgSetupService.healthCheck,
        getSummary: function() {
          return OrgCrudService.getModuleSummary();
        }
      });
    } catch (e) {
      console.log('[OrgSetupService] ModuleRegistry registration deferred');
    }
  }
})();
