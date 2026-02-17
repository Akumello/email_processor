/**
 * Teams CRUD Service - Create, Read, Update, Delete operations for Teams
 * Manages the TEAM_MAPPINGS sheet for team/task/workstream configuration
 * 
 * Sheet Columns: Contract, Task ID, Task Name, Team ID, Team Name, Is Active,
 *                Color, Description, Default SLA Threshold, Notify On Escalation, Display Order
 * 
 * Also manages vacant position rows in a VACANT_POSITIONS sheet for positions
 * that never had a person assigned.
 * 
 * DATA SOURCES:
 * - Task Names: Canonical source for task display names (e.g., "Task 1 - Program Management")
 * - Colors: Task/team color assignments for org chart rendering
 * - Thresholds: Task-level SLA threshold defaults
 * - Management emails: Now derived from Team List personnel data via OrgCrudService
 */

const TeamsCrudService = (function() {
  'use strict';
  
  // Configuration
  const SHEET_NAME = 'Team Mappings';
  const VACANT_SHEET_NAME = 'Vacant Positions';
  const CACHE_KEY_ALL_TEAMS = 'teams:all';
  const CACHE_KEY_VACANTS = 'teams:vacants';
  const CACHE_TTL = 300; // 5 minutes
  
  // Headers for the TEAM_MAPPINGS sheet (extended with metadata)
  const HEADERS = [
    'Contract',               // Contract name (e.g., "SQuAT", "Forward")
    'Task ID',                // Task identifier (e.g., "TASK-001")
    'Task Name',              // Canonical task display name (e.g., "Task 1 - Program Management")
    'Team ID',                // Unique team identifier (e.g., "TEAM-001") - workstream for Forward
    'Team Name',              // Human-readable team name - empty for Forward
    'Is Active',              // Whether team/task is active (true/false)
    'Color',                  // Hex color for org chart rendering (e.g., "#9b59b6")
    'Description',            // Team/task description
    'Default SLA Threshold',  // Task-level default SLA threshold percentage
    'Notify On Escalation',   // Task-level escalation notification flag (true/false)
    'Display Order'           // Sort weight for ordering tasks/teams in the org chart
  ];
  
  // Headers for the VACANT_POSITIONS sheet
  const VACANT_HEADERS = [
    'Vacant ID',              // Synthetic ID (e.g., "VAC-TASK001-1")
    'Contract',               // Contract name
    'Task ID',                // Parent task
    'Team ID',                // Parent team (optional)
    'Title',                  // Position title (e.g., "Junior Analyst")
    'Supervisor UPID',        // UPID of the supervisor this vacancy reports to
    'Target Hire Date',       // Expected fill date
    'Requirements',           // Position requirements
    'Is Active'               // Whether this vacancy is still open
  ];
  
  /**
   * Get the spreadsheet (uses SLA data store since teams are shared)
   * @private
   */
  function _getSpreadsheet() {
    if (typeof DataStoreRegistry !== 'undefined') {
      return DataStoreRegistry.getSpreadsheet('sla');
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  /**
   * Get the TEAMS sheet, create if not exists
   * @private
   */
  function _getSheet() {
    const ss = _getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.log('[TeamsCrudService] Sheet not found, creating: ' + SHEET_NAME);
      sheet = ss.insertSheet(SHEET_NAME);
      
      // Set up headers
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#4F46E5')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  }
  
  /**
   * Invalidate cache
   * @private
   */
  function _invalidateCache() {
    if (typeof AppCache !== 'undefined') {
      AppCache.invalidate(CACHE_KEY_ALL_TEAMS);
      console.log('[TeamsCrudService] Cache invalidated');
    }
  }
  
  /**
   * Generate next Team ID
   * @private
   * @returns {string} Unique team ID (e.g., "TEAM-013")
   */
  function _generateTeamId() {
    const allTeams = getAllTeams();
    let maxNum = 0;
    
    allTeams.forEach(team => {
      const match = team.teamId.match(/TEAM-(\d+)/);
      if (match) {
        const num = parseInt(match[1]);
        if (num > maxNum) maxNum = num;
      }
    });
    
    return `TEAM-${String(maxNum + 1).padStart(3, '0')}`;
  }
  
  /**
   * Get all teams from the sheet
   * @returns {Array<Object>} Array of team objects
   */
  function getAllTeams() {
    try {
      // Check cache first
      if (typeof AppCache !== 'undefined') {
        const cached = AppCache.get(CACHE_KEY_ALL_TEAMS);
        if (cached) {
          console.log('[TeamsCrudService] Returning cached teams');
          return cached;
        }
      }
      
      const sheet = _getSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length < 2) {
        return [];
      }
      
      const headers = data[0];
      const teams = [];
      
      // Build column index map
      const colIdx = {};
      headers.forEach((header, index) => {
        colIdx[header] = index;
      });
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip empty rows (need either Team ID for SQuAT or Task ID for Forward)
        if (!row[colIdx['Team ID']] && !row[colIdx['Task ID']]) continue;
        
        const team = {
          contract: row[colIdx['Contract']] || 'SQuAT',
          task: row[colIdx['Task ID']] || '',
          taskName: row[colIdx['Task Name']] || '',
          teamId: row[colIdx['Team ID']] || '',
          teamName: row[colIdx['Team Name']] || '',
          isActive: row[colIdx['Is Active']] !== false && row[colIdx['Is Active']] !== 'FALSE',
          color: row[colIdx['Color']] || '',
          description: row[colIdx['Description']] || '',
          defaultSlaThreshold: row[colIdx['Default SLA Threshold']] || '',
          notifyOnEscalation: row[colIdx['Notify On Escalation']] === true || row[colIdx['Notify On Escalation']] === 'TRUE',
          displayOrder: row[colIdx['Display Order']] || 0,
          rowIndex: i + 1 // 1-based row number for updates
        };
        
        teams.push(team);
      }
      
      // Cache results
      if (typeof AppCache !== 'undefined') {
        AppCache.set(CACHE_KEY_ALL_TEAMS, teams, CACHE_TTL);
      }
      
      console.log('[TeamsCrudService] Loaded ' + teams.length + ' teams');
      return teams;
      
    } catch (error) {
      console.error('[TeamsCrudService] Error getting teams:', error);
      return [];
    }
  }
  
  /**
   * Get team by ID
   * @param {string} teamId - Team ID (e.g., "TEAM-001")
   * @returns {Object|null} Team object or null if not found
   */
  function getTeamById(teamId) {
    const teams = getAllTeams();
    return teams.find(t => t.teamId === teamId) || null;
  }
  
  /**
   * Get teams by task
   * @param {string} task - Task identifier (e.g., "Task 1")
   * @returns {Array<Object>} Array of team objects for the task
   */
  function getTeamsByTask(task) {
    const teams = getAllTeams();
    return teams.filter(t => t.task === task && t.isActive);
  }
  
  /**
   * Get all unique tasks
   * @returns {Array<Object>} Array of {task, taskName} objects
   */
  function getAllTasks() {
    const teams = getAllTeams();
    const taskMap = new Map();
    
    teams.forEach(team => {
      if (team.task && !taskMap.has(team.task)) {
        taskMap.set(team.task, {
          task: team.task,
          taskName: team.taskName || team.task
        });
      }
    });
    
    // Sort by task number
    return Array.from(taskMap.values()).sort((a, b) => {
      const numA = parseInt(a.task.replace('Task ', '')) || 0;
      const numB = parseInt(b.task.replace('Task ', '')) || 0;
      return numA - numB;
    });
  }
  
  /**
   * Create a new team
   * @param {Object} teamData - Team data
   * @param {string} teamData.task - Task identifier
   * @param {string} teamData.taskName - Task display name
   * @param {string} teamData.teamName - Team name
   * @param {string} [teamData.managerEmail] - Manager email
   * @param {boolean} [teamData.isActive=true] - Active status
   * @returns {Object} Result with success status and team data
   */
  function createTeam(teamData) {
    try {
      console.log('[TeamsCrudService] Creating team:', teamData.teamName);
      
      const sheet = _getSheet();
      const teamId = teamData.teamId || _generateTeamId();
      
      // Check if team ID already exists
      const existing = getTeamById(teamId);
      if (existing) {
        return {
          success: false,
          error: 'Team ID already exists: ' + teamId
        };
      }
      
      const rowData = [
        teamData.contract || 'SQuAT',
        teamData.task || '',
        teamData.taskName || '',
        teamId,
        teamData.teamName || '',
        teamData.isActive !== false,
        teamData.color || '',
        teamData.description || '',
        teamData.defaultSlaThreshold || '',
        teamData.notifyOnEscalation || false,
        teamData.displayOrder || 0
      ];
      
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
      
      _invalidateCache();
      
      console.log('[TeamsCrudService] Created team:', teamId);
      
      return {
        success: true,
        teamId: teamId,
        message: 'Team created successfully'
      };
      
    } catch (error) {
      console.error('[TeamsCrudService] Error creating team:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Update an existing team
   * @param {string} teamId - Team ID to update
   * @param {Object} updates - Fields to update
   * @returns {Object} Result with success status
   */
  function updateTeam(teamId, updates) {
    try {
      console.log('[TeamsCrudService] Updating team:', teamId);
      
      const team = getTeamById(teamId);
      if (!team) {
        return {
          success: false,
          error: 'Team not found: ' + teamId
        };
      }
      
      const sheet = _getSheet();
      const rowIndex = team.rowIndex;
      
      // Merge updates with existing data
      const updatedTeam = {
        contract: updates.contract !== undefined ? updates.contract : team.contract,
        task: updates.task !== undefined ? updates.task : team.task,
        taskName: updates.taskName !== undefined ? updates.taskName : team.taskName,
        teamId: teamId, // Cannot change team ID
        teamName: updates.teamName !== undefined ? updates.teamName : team.teamName,
        isActive: updates.isActive !== undefined ? updates.isActive : team.isActive,
        color: updates.color !== undefined ? updates.color : team.color,
        description: updates.description !== undefined ? updates.description : team.description,
        defaultSlaThreshold: updates.defaultSlaThreshold !== undefined ? updates.defaultSlaThreshold : team.defaultSlaThreshold,
        notifyOnEscalation: updates.notifyOnEscalation !== undefined ? updates.notifyOnEscalation : team.notifyOnEscalation,
        displayOrder: updates.displayOrder !== undefined ? updates.displayOrder : team.displayOrder
      };
      
      const rowData = [
        updatedTeam.contract,
        updatedTeam.task,
        updatedTeam.taskName,
        updatedTeam.teamId,
        updatedTeam.teamName,
        updatedTeam.isActive,
        updatedTeam.color,
        updatedTeam.description,
        updatedTeam.defaultSlaThreshold,
        updatedTeam.notifyOnEscalation,
        updatedTeam.displayOrder
      ];
      
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      
      _invalidateCache();
      
      console.log('[TeamsCrudService] Updated team:', teamId);
      
      return {
        success: true,
        message: 'Team updated successfully'
      };
      
    } catch (error) {
      console.error('[TeamsCrudService] Error updating team:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Delete a team (soft delete - sets isActive to false)
   * @param {string} teamId - Team ID to delete
   * @param {boolean} [hardDelete=false] - If true, permanently removes the row
   * @returns {Object} Result with success status
   */
  function deleteTeam(teamId, hardDelete = false) {
    try {
      console.log('[TeamsCrudService] Deleting team:', teamId, hardDelete ? '(hard)' : '(soft)');
      
      const team = getTeamById(teamId);
      if (!team) {
        return {
          success: false,
          error: 'Team not found: ' + teamId
        };
      }
      
      const sheet = _getSheet();
      
      if (hardDelete) {
        // Permanently delete the row
        sheet.deleteRow(team.rowIndex);
      } else {
        // Soft delete - set isActive to false
        const isActiveCol = HEADERS.indexOf('Is Active') + 1;
        sheet.getRange(team.rowIndex, isActiveCol).setValue(false);
      }
      
      _invalidateCache();
      
      console.log('[TeamsCrudService] Deleted team:', teamId);
      
      return {
        success: true,
        message: hardDelete ? 'Team permanently deleted' : 'Team deactivated'
      };
      
    } catch (error) {
      console.error('[TeamsCrudService] Error deleting team:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Reactivate a deactivated team
   * @param {string} teamId - Team ID to reactivate
   * @returns {Object} Result with success status
   */
  function reactivateTeam(teamId) {
    return updateTeam(teamId, { isActive: true });
  }
  
  /**
   * Bulk create teams
   * @param {Array<Object>} teamsData - Array of team data objects
   * @returns {Object} Result with success counts
   */
  function bulkCreateTeams(teamsData) {
    try {
      console.log('[TeamsCrudService] Bulk creating ' + teamsData.length + ' teams');
      
      let successCount = 0;
      let failCount = 0;
      const errors = [];
      
      teamsData.forEach((teamData, index) => {
        const result = createTeam(teamData);
        if (result.success) {
          successCount++;
        } else {
          failCount++;
          errors.push({ index, error: result.error });
        }
      });
      
      return {
        success: failCount === 0,
        created: successCount,
        failed: failCount,
        errors: errors,
        message: `Created ${successCount} teams, ${failCount} failed`
      };
      
    } catch (error) {
      console.error('[TeamsCrudService] Error in bulk create:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Get module summary for health checks
   * @returns {Object} Summary statistics
   */
  function getModuleSummary() {
    const teams = getAllTeams();
    const activeTeams = teams.filter(t => t.isActive);
    const tasks = getAllTasks();
    
    return {
      totalTeams: teams.length,
      activeTeams: activeTeams.length,
      inactiveTeams: teams.length - activeTeams.length,
      totalTasks: tasks.length
    };
  }
  
  /**
   * Clear cache
   * @returns {Object} Result
   */
  function clearCache() {
    _invalidateCache();
    return { success: true, message: 'Cache cleared' };
  }
  
  // ============================================================================
  // TASK METADATA (colors, names, thresholds)
  // ============================================================================
  
  /**
   * Get task metadata from Team Mappings (colors, names, thresholds)
   * Aggregates at the task level from all team rows for that task
   * @returns {Object} Map of task IDs to metadata
   */
  function getTaskMetadata() {
    const teams = getAllTeams();
    const taskMeta = {};
    
    teams.forEach(team => {
      if (!team.task) return;
      
      if (!taskMeta[team.task]) {
        taskMeta[team.task] = {
          taskId: team.task,
          taskName: team.taskName || team.task,
          contract: team.contract || '',
          color: team.color || '',
          description: team.description || '',
          defaultSlaThreshold: team.defaultSlaThreshold || '',
          notifyOnEscalation: team.notifyOnEscalation || false,
          displayOrder: team.displayOrder || 0
        };
      } else {
        // Prefer non-empty values from any row for this task
        const meta = taskMeta[team.task];
        if (!meta.taskName && team.taskName) meta.taskName = team.taskName;
        if (!meta.color && team.color) meta.color = team.color;
        if (!meta.description && team.description) meta.description = team.description;
        if (!meta.defaultSlaThreshold && team.defaultSlaThreshold) meta.defaultSlaThreshold = team.defaultSlaThreshold;
        if (!meta.notifyOnEscalation && team.notifyOnEscalation) meta.notifyOnEscalation = team.notifyOnEscalation;
      }
    });
    
    return taskMeta;
  }
  
  /**
   * Get task colors from Team Mappings
   * @returns {Object} Map of task IDs to hex colors { 'TASK-001': '#9b59b6', ... }
   */
  function getTaskColors() {
    const taskMeta = getTaskMetadata();
    const colors = { 'default': '#95a5a6' };
    
    Object.entries(taskMeta).forEach(([taskId, meta]) => {
      if (meta.color) {
        colors[taskId] = meta.color;
      }
    });
    
    return colors;
  }
  
  /**
   * Get task friendly names from Team Mappings
   * @returns {Object} Map of task IDs to display names { 'TASK-001': 'Task 1 - Program Management', ... }
   */
  function getTaskFriendlyNames() {
    const taskMeta = getTaskMetadata();
    const names = {};
    
    Object.entries(taskMeta).forEach(([taskId, meta]) => {
      names[taskId] = meta.taskName || taskId;
    });
    
    return names;
  }
  
  // ============================================================================
  // VACANT POSITIONS (positions that never had a person)
  // ============================================================================
  
  /**
   * Get or create the Vacant Positions sheet
   * @private
   */
  function _getVacantSheet() {
    const ss = _getSpreadsheet();
    let sheet = ss.getSheetByName(VACANT_SHEET_NAME);
    
    if (!sheet) {
      console.log('[TeamsCrudService] Vacant Positions sheet not found, creating...');
      sheet = ss.insertSheet(VACANT_SHEET_NAME);
      sheet.getRange(1, 1, 1, VACANT_HEADERS.length).setValues([VACANT_HEADERS]);
      sheet.getRange(1, 1, 1, VACANT_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#E74C3C')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  }
  
  /**
   * Get all vacant positions
   * @returns {Array<Object>} Array of vacant position objects
   */
  function getAllVacantPositions() {
    try {
      if (typeof AppCache !== 'undefined') {
        const cached = AppCache.get(CACHE_KEY_VACANTS);
        if (cached) return cached;
      }
      
      const sheet = _getVacantSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length < 2) return [];
      
      const headers = data[0];
      const colIdx = {};
      headers.forEach((h, i) => colIdx[h] = i);
      
      const vacants = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[colIdx['Vacant ID']]) continue;
        
        const isActive = row[colIdx['Is Active']];
        if (isActive === false || isActive === 'FALSE') continue;
        
        vacants.push({
          vacantId: String(row[colIdx['Vacant ID']]),
          contract: String(row[colIdx['Contract']] || ''),
          task: String(row[colIdx['Task ID']] || ''),
          team: String(row[colIdx['Team ID']] || ''),
          title: String(row[colIdx['Title']] || 'Vacant Position'),
          supervisorUpid: String(row[colIdx['Supervisor UPID']] || ''),
          targetHireDate: row[colIdx['Target Hire Date']] instanceof Date
            ? row[colIdx['Target Hire Date']].toISOString()
            : String(row[colIdx['Target Hire Date']] || ''),
          requirements: String(row[colIdx['Requirements']] || ''),
          isActive: true,
          rowIndex: i + 1
        });
      }
      
      if (typeof AppCache !== 'undefined') {
        AppCache.set(CACHE_KEY_VACANTS, vacants, CACHE_TTL);
      }
      
      return vacants;
    } catch (error) {
      console.error('[TeamsCrudService] Error getting vacant positions:', error);
      return [];
    }
  }
  
  /**
   * Create a new vacant position
   * @param {Object} vacantData - Vacant position data
   * @returns {Object} Result with success status
   */
  function createVacantPosition(vacantData) {
    try {
      const sheet = _getVacantSheet();
      
      // Generate vacant ID
      const existing = getAllVacantPositions();
      const taskSuffix = (vacantData.task || 'UNKNOWN').replace('TASK-', '');
      const count = existing.filter(v => v.task === vacantData.task).length + 1;
      const vacantId = vacantData.vacantId || `VAC-${taskSuffix}-${count}`;
      
      const rowData = [
        vacantId,
        vacantData.contract || '',
        vacantData.task || '',
        vacantData.team || '',
        vacantData.title || 'Vacant Position',
        vacantData.supervisorUpid || '',
        vacantData.targetHireDate || '',
        vacantData.requirements || '',
        true
      ];
      
      sheet.appendRow(rowData);
      _invalidateCache();
      if (typeof AppCache !== 'undefined') AppCache.invalidate(CACHE_KEY_VACANTS);
      
      return { success: true, vacantId: vacantId };
    } catch (error) {
      console.error('[TeamsCrudService] Error creating vacant position:', error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Delete (deactivate) a vacant position
   * @param {string} vacantId - Vacant position ID
   * @returns {Object} Result with success status
   */
  function deleteVacantPosition(vacantId) {
    try {
      const vacants = getAllVacantPositions();
      const vacant = vacants.find(v => v.vacantId === vacantId);
      if (!vacant) return { success: false, error: 'Vacant position not found: ' + vacantId };
      
      const sheet = _getVacantSheet();
      const isActiveCol = VACANT_HEADERS.indexOf('Is Active') + 1;
      sheet.getRange(vacant.rowIndex, isActiveCol).setValue(false);
      
      _invalidateCache();
      if (typeof AppCache !== 'undefined') AppCache.invalidate(CACHE_KEY_VACANTS);
      
      return { success: true };
    } catch (error) {
      console.error('[TeamsCrudService] Error deleting vacant position:', error);
      return { success: false, error: error.message };
    }
  }
  
  // Public API
  return {
    getAllTeams: getAllTeams,
    getTeamById: getTeamById,
    getTeamsByTask: getTeamsByTask,
    getAllTasks: getAllTasks,
    createTeam: createTeam,
    updateTeam: updateTeam,
    deleteTeam: deleteTeam,
    reactivateTeam: reactivateTeam,
    bulkCreateTeams: bulkCreateTeams,
    getModuleSummary: getModuleSummary,
    clearCache: clearCache,
    // Task metadata (colors, names, thresholds)
    getTaskMetadata: getTaskMetadata,
    getTaskColors: getTaskColors,
    getTaskFriendlyNames: getTaskFriendlyNames,
    // Vacant positions
    getAllVacantPositions: getAllVacantPositions,
    createVacantPosition: createVacantPosition,
    deleteVacantPosition: deleteVacantPosition,
    // Constants
    HEADERS: HEADERS,
    VACANT_HEADERS: VACANT_HEADERS
  };
})();
