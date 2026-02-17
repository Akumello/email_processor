/**
 * Organization CRUD Service
 * Handles data operations for the organization module
 * 
 * DATA SOURCES:
 * - External "Team List" sheet (via DataStoreRegistry 'teamlist') — personnel roster
 *   Primary key: UPID (Unique Personnel ID, format: CPC-HID, e.g. 310-003)
 *     CPC = 3 digits (Level + Task + Role), HID = 3-digit sequential number
 *   Parent pointer: Supervisor UPID
 * - "Team Mappings" sheet (via TeamsCrudService) — task/team structure & metadata
 * - "Vacant Positions" sheet (via TeamsCrudService) — unfilled positions
 * 
 * ARCHITECTURE:
 * - Team List contains ONLY people rows (no structural task/team nodes)
 * - Structural nodes (task, team, hidden root) are derived at read time
 *   from Team Mappings data and personnel groupings
 * - The tree is built using UPID-based parent pointers for people
 *   and synthetic IDs for structural nodes (e.g., "task:TASK-001", "team:TEAM-001")
 * - Departed personnel with "Active In Org" = TRUE render as vacant positions
 * 
 * Caching Strategy:
 * - getAllData: Cached for 5 minutes (invalidated on add/update/delete)
 * - External Team List raw read cached separately for longer TTL
 */

const OrgCrudService = (function() {
  'use strict';

  const TEAM_LIST_SHEET_NAME = 'Team List';
  const DATASTORE_MODULE = 'teamlist';
  
  // Team List column headers (external sheet)
  const TEAM_LIST_HEADERS = [
    'Employee Code',
    'Company',
    'Contract',
    'Task',
    'Primary Workstream',
    'Secondary Workstream',
    'First Name',
    'Last Name',
    'Email',
    'Primary Role',
    'Secondary Role',
    'Primary Role Start Date',
    'Contract Personnel Code (CPC)',
    'Heirarchy Identifier (HID)',
    'Unique Personnel ID (UPID)',
    'Supervisor Email',
    'Supervisor UPID',
    'Portfolio Leadership?',
    'Profile Picture',
    'EOD',
    'Personnel Contract Status',
    'Departure Date',
    'Departure Meeting Date',
    'Contract LCAT',
    'Location (City, ST)',
    'Tenure (Days)',
    'Node Type',
    'Active In Org'
  ];
  
  // Exported HEADERS for backward compatibility (represents the unified record shape)
  const HEADERS = [
    'id', 'parentId', 'company', 'contract', 'task', 'team',
    'name', 'email', 'title', 'type', 'primaryRole', 'secondaryRoles',
    'cpc', 'hid', 'upid', 'profilePicture', 'eod', 'personnelContractStatus',
    'notifyOnEscalation', 'defaultSlaThreshold', 'color', 'description',
    'targetHireDate', 'requirements', 'active'
  ];
  
  const CACHE_KEY_ALL_DATA = 'org:all_data';
  const CACHE_KEY_TEAM_LIST_RAW = 'org:teamlist_raw';
  const CACHE_TTL = 300; // 5 minutes
  const CACHE_TTL_RAW = 600; // 10 minutes for raw external read
  
  // Structural ID prefixes
  const ID_PREFIX_TASK = 'task:';
  const ID_PREFIX_TEAM = 'team:';
  const ID_ROOT = 'root';
  
  // ============================================================================
  // PRIVATE: External Sheet Access
  // ============================================================================
  
  /**
   * Get the external Team List sheet via DataStoreRegistry
   * Falls back to bound spreadsheet if not configured
   * @private
   * @returns {Sheet} Team List sheet
   */
  function _getTeamListSheet() {
    if (typeof DataStoreRegistry !== 'undefined') {
      const sheet = DataStoreRegistry.getSheet(DATASTORE_MODULE, TEAM_LIST_SHEET_NAME);
      if (sheet) return sheet;
    }
    // Fallback to bound spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TEAM_LIST_SHEET_NAME);
    if (!sheet) {
      throw new Error('[OrgCrudService] Team List sheet not found. Register the external spreadsheet via DataStoreRegistry.setSpreadsheetId("teamlist", "SPREADSHEET_ID")');
    }
    return sheet;
  }
  
  /**
   * Invalidate cache (call after any write operation)
   * Also clears TeamConfigService cache since task friendly names and management emails
   * are derived from personnel data
   * @private
   */
  function _invalidateCache() {
    if (typeof AppCache !== 'undefined') {
      AppCache.invalidate(CACHE_KEY_ALL_DATA);
      AppCache.invalidate(CACHE_KEY_TEAM_LIST_RAW);
      console.log('[OrgCrudService] Cache invalidated');
    }
    if (typeof TeamConfigService !== 'undefined') {
      TeamConfigService.clearCache();
      console.log('[OrgCrudService] TeamConfigService cache also cleared');
    }
  }
  
  // ============================================================================
  // PRIVATE: Read Raw Personnel Data from Team List
  // ============================================================================
  
  /**
   * Read raw personnel records from the external Team List sheet
   * @private
   * @returns {Array<Object>} Array of raw personnel records
   */
  function _readTeamListRaw() {
    // Check raw cache
    if (typeof AppCache !== 'undefined') {
      const cached = AppCache.get(CACHE_KEY_TEAM_LIST_RAW);
      if (cached) {
        console.log('[OrgCrudService] Returning cached raw Team List');
        return cached;
      }
    }
    
    const sheet = _getTeamListSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) return [];
    
    const headers = data[0];
    const colMap = {};
    headers.forEach((header, index) => {
      colMap[String(header).trim()] = index;
    });
    
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip completely empty rows
      const upid = row[colMap['Unique Personnel ID (UPID)']];
      const email = row[colMap['Email']];
      if (!upid && !email) continue;
      
      const personnelStatus = String(row[colMap['Personnel Contract Status']] || '').trim();
      const activeInOrg = row[colMap['Active In Org']];
      
      // Determine if this row should be excluded entirely
      // Departed + Active In Org explicitly FALSE → excluded
      if (personnelStatus === 'Departed' && (activeInOrg === false || activeInOrg === 'FALSE' || activeInOrg === 'false')) {
        continue;
      }
      
      const firstName = String(row[colMap['First Name']] || '').trim();
      const lastName = String(row[colMap['Last Name']] || '').trim();
      const fullName = [firstName, lastName].filter(Boolean).join(' ');
      
      // Determine node type
      let nodeType = String(row[colMap['Node Type']] || '').trim().toLowerCase();
      if (!nodeType || nodeType === '') {
        // Auto-derive from CPC if possible
        const cpc = String(row[colMap['Contract Personnel Code (CPC)']] || '');
        if (cpc && cpc.length >= 1) {
          const levelDigit = cpc.charAt(0);
          const cpcMap = (typeof CONFIG !== 'undefined' && CONFIG.ORG) ? CONFIG.ORG.CPC_LEVEL_MAP : { '1': 'director', '2': 'deputy', '3': 'lead', '4': 'person' };
          nodeType = cpcMap[levelDigit] || 'person';
        } else {
          nodeType = 'person';
        }
        // Override with Portfolio Leadership flag
        const isLeadership = row[colMap['Portfolio Leadership?']];
        if (isLeadership === true || isLeadership === 'TRUE' || isLeadership === 'Yes') {
          if (nodeType === 'person') nodeType = 'director';
        }
      }
      
      // Vacancy detection: departed but position still active in org
      const isDeparted = personnelStatus === 'Departed';
      if (isDeparted && activeInOrg !== false && activeInOrg !== 'FALSE' && activeInOrg !== 'false') {
        nodeType = 'vacant';
      }
      
      const record = {
        // Identity
        employeeCode: String(row[colMap['Employee Code']] || ''),
        upid: String(upid || ''),
        cpc: String(row[colMap['Contract Personnel Code (CPC)']] || ''),
        hid: String(row[colMap['Heirarchy Identifier (HID)']] || ''),
        
        // Hierarchy
        supervisorEmail: String(row[colMap['Supervisor Email']] || '').trim(),
        supervisorUpid: String(row[colMap['Supervisor UPID']] || '').trim(),
        
        // Organization
        company: String(row[colMap['Company']] || ''),
        contract: String(row[colMap['Contract']] || ''),
        task: String(row[colMap['Task']] || ''),
        primaryWorkstream: String(row[colMap['Primary Workstream']] || ''),
        secondaryWorkstream: String(row[colMap['Secondary Workstream']] || ''),
        
        // Person info
        firstName: firstName,
        lastName: lastName,
        name: fullName,
        email: String(email || '').trim(),
        primaryRole: String(row[colMap['Primary Role']] || ''),
        secondaryRole: String(row[colMap['Secondary Role']] || ''),
        primaryRoleStartDate: row[colMap['Primary Role Start Date']] instanceof Date
          ? row[colMap['Primary Role Start Date']].toISOString()
          : String(row[colMap['Primary Role Start Date']] || ''),
        profilePicture: String(row[colMap['Profile Picture']] || ''),
        eod: row[colMap['EOD']] instanceof Date
          ? row[colMap['EOD']].toISOString()
          : String(row[colMap['EOD']] || ''),
        personnelContractStatus: personnelStatus,
        departureDate: row[colMap['Departure Date']] instanceof Date
          ? row[colMap['Departure Date']].toISOString()
          : String(row[colMap['Departure Date']] || ''),
        departureMeetingDate: row[colMap['Departure Meeting Date']] instanceof Date
          ? row[colMap['Departure Meeting Date']].toISOString()
          : String(row[colMap['Departure Meeting Date']] || ''),
        contractLcat: String(row[colMap['Contract LCAT']] || ''),
        location: String(row[colMap['Location (City, ST)']] || ''),
        tenure: row[colMap['Tenure (Days)']] || '',
        
        // Classification
        nodeType: nodeType,
        portfolioLeadership: row[colMap['Portfolio Leadership?']] === true || row[colMap['Portfolio Leadership?']] === 'TRUE' || row[colMap['Portfolio Leadership?']] === 'Yes',
        activeInOrg: activeInOrg !== false && activeInOrg !== 'FALSE' && activeInOrg !== 'false',
        
        // Sheet row for writes
        _rowIndex: i + 1
      };
      
      records.push(record);
    }
    
    // Cache raw data
    if (typeof AppCache !== 'undefined') {
      AppCache.set(CACHE_KEY_TEAM_LIST_RAW, records, CACHE_TTL_RAW);
      console.log('[OrgCrudService] Raw Team List cached:', records.length, 'records');
    }
    
    return records;
  }
  
  // ============================================================================
  // PRIVATE: Tree Derivation (structural nodes from flat data)
  // ============================================================================
  
  /**
   * Derive structural task and team nodes from Team Mappings + personnel data
   * @private
   * @param {Array<Object>} personnel - Raw personnel records
   * @returns {Object} { taskNodes, teamNodes, rootNode, taskMeta }
   */
  function _deriveStructuralNodes(personnel) {
    // Get task/team metadata from TeamsCrudService
    const teams = (typeof TeamsCrudService !== 'undefined') ? TeamsCrudService.getAllTeams() : [];
    const taskMeta = (typeof TeamsCrudService !== 'undefined') ? TeamsCrudService.getTaskMetadata() : {};
    
    // Find unique tasks from both Team Mappings and personnel
    const taskIds = new Set();
    teams.forEach(t => { if (t.task && t.isActive) taskIds.add(t.task); });
    personnel.forEach(p => { if (p.task) taskIds.add(p.task); });
    
    // Find unique teams from Team Mappings
    const teamIds = new Set();
    const teamToTask = {};
    teams.forEach(t => {
      if (t.teamId && t.isActive) {
        teamIds.add(t.teamId);
        teamToTask[t.teamId] = t.task;
      }
    });
    
    // Find directors/deputies per contract to determine task node parents
    const contractDirectors = {};
    personnel.forEach(p => {
      if (p.nodeType === 'director' && p.contract && p.upid) {
        contractDirectors[p.contract] = p.upid;
      }
    });
    
    // Build task nodes
    const taskNodes = [];
    taskIds.forEach(taskId => {
      const meta = taskMeta[taskId] || {};
      const contract = meta.contract || '';
      const directorUpid = contractDirectors[contract] || null;
      
      taskNodes.push({
        id: ID_PREFIX_TASK + taskId,
        // Task node parent: director's UPID for the same contract, or root
        parentId: directorUpid || ID_ROOT,
        company: '',
        contract: contract,
        task: taskId,
        team: null,
        name: meta.taskName || taskId,
        email: '',
        title: meta.description || '',
        type: 'task',
        active: true,
        notifyOnEscalation: meta.notifyOnEscalation || false,
        defaultSlaThreshold: meta.defaultSlaThreshold || '',
        color: meta.color || '',
        description: meta.description || '',
        displayOrder: meta.displayOrder || 0,
        _isStructural: true
      });
    });
    
    // Build team nodes
    const teamNodes = [];
    teams.forEach(t => {
      if (!t.teamId || !t.isActive) return;
      
      teamNodes.push({
        id: ID_PREFIX_TEAM + t.teamId,
        parentId: ID_PREFIX_TASK + t.task,
        company: '',
        contract: t.contract || '',
        task: t.task,
        team: t.teamId,
        name: t.teamName || t.teamId,
        email: '',
        title: '',
        type: 'team',
        active: true,
        color: t.color || '',
        description: t.description || '',
        _isStructural: true
      });
    });
    
    // Build root node (hidden)
    const rootNode = {
      id: ID_ROOT,
      parentId: null,
      company: '',
      contract: '',
      task: null,
      team: null,
      name: '',
      email: '',
      title: '',
      type: 'hidden',
      active: true,
      _isStructural: true
    };
    
    return { taskNodes, teamNodes, rootNode, taskMeta, teamToTask };
  }
  
  /**
   * Determine the parentId for a personnel record in the tree
   * Resolves Supervisor UPID → UPID for person-to-person links,
   * or assigns to team/task structural node if no supervisor found
   * @private
   * @param {Object} person - Personnel record
   * @param {Map<string, Object>} upidMap - UPID → personnel record map
   * @param {Object} teamToTask - Team ID → Task ID map
   * @param {Set<string>} teamNodeIds - Set of team node IDs that exist
   * @param {Set<string>} taskNodeIds - Set of task node IDs that exist
   * @returns {string} Parent node ID
   */
  function _resolveParentId(person, upidMap, teamToTask, teamNodeIds, taskNodeIds) {
    // 1. Directors/Deputies report to root
    if (person.nodeType === 'director' || person.nodeType === 'deputy') {
      return ID_ROOT;
    }
    
    // 2. If person has a task, try to place under their team or task node first.
    //    Within-task supervisor links are resolved below (step 3).
    if (person.task) {
      // 2a. If person has a team (workstream), parent is the team node
      if (person.primaryWorkstream) {
        const teams = (typeof TeamsCrudService !== 'undefined') ? TeamsCrudService.getAllTeams() : [];
        const matchedTeam = teams.find(t =>
          t.teamName === person.primaryWorkstream && t.task === person.task && t.isActive
        );
        if (matchedTeam && matchedTeam.teamId) {
          const teamNodeId = ID_PREFIX_TEAM + matchedTeam.teamId;
          if (teamNodeIds.has(teamNodeId)) {
            return teamNodeId;
          }
        }
      }
      
      // 2b. If supervisor is on the same task, link to them (within-task hierarchy)
      if (person.supervisorUpid && upidMap.has(person.supervisorUpid)) {
        const supervisor = upidMap.get(person.supervisorUpid);
        if (supervisor.task === person.task) {
          return person.supervisorUpid;
        }
      }
      
      // 2c. Fall back to task node
      const taskNodeId = ID_PREFIX_TASK + person.task;
      if (taskNodeIds.has(taskNodeId)) {
        return taskNodeId;
      }
    }
    
    // 3. No task — use direct supervisor link if available
    if (person.supervisorUpid && upidMap.has(person.supervisorUpid)) {
      return person.supervisorUpid;
    }
    
    // 4. Last resort: root
    return ID_ROOT;
  }
  
  // ============================================================================
  // PUBLIC: Core Read Operations
  // ============================================================================
  
  /**
   * Get all org chart data (with caching)
   * Joins external Team List personnel + Team Mappings structure + Vacant Positions
   * Returns a unified array with the same record shape as the old Organization sheet
   * @returns {Array<Object>} Array of node objects (people + structural + vacant)
   */
  function getAllData() {
    // Try to get from cache first
    if (typeof AppCache !== 'undefined') {
      const cached = AppCache.get(CACHE_KEY_ALL_DATA);
      if (cached) {
        console.log('[OrgCrudService] Returning cached org data');
        return cached;
      }
    }
    
    try {
      // 1. Read raw personnel from external Team List
      const personnel = _readTeamListRaw();
      
      // 2. Derive structural nodes (tasks, teams, root)
      const { taskNodes, teamNodes, rootNode, taskMeta, teamToTask } = _deriveStructuralNodes(personnel);
      
      // 3. Build lookup maps for parent resolution
      const upidMap = new Map();
      personnel.forEach(p => {
        if (p.upid) upidMap.set(p.upid, p);
      });
      
      const teamNodeIds = new Set(teamNodes.map(t => t.id));
      const taskNodeIds = new Set(taskNodes.map(t => t.id));
      
      // 4. Transform personnel → unified record shape
      const personRecords = personnel.map(p => {
        const parentId = _resolveParentId(p, upidMap, teamToTask, teamNodeIds, taskNodeIds);
        
        const record = {
          // Core fields — id is UPID (the stable position identifier)
          id: p.upid || p.email || p.employeeCode,
          parentId: parentId,
          company: p.company,
          contract: p.contract,
          task: p.task || null,
          team: p.primaryWorkstream || null,
          name: p.nodeType === 'vacant' ? ('VACANT - ' + (p.name || 'Position')) : p.name,
          email: p.email,
          title: p.primaryRole || '',
          type: p.nodeType,
          active: true,
          
          // Personnel detail fields
          employeeCode: p.employeeCode,
          firstName: p.firstName,
          lastName: p.lastName,
          upid: p.upid,
          cpc: p.cpc,
          hid: p.hid,
          supervisorUpid: p.supervisorUpid,
          supervisorEmail: p.supervisorEmail,
          primaryWorkstream: p.primaryWorkstream,
          secondaryWorkstream: p.secondaryWorkstream
        };
        
        // Conditionally add non-empty fields
        if (p.primaryRole) record.primaryRole = p.primaryRole;
        if (p.secondaryRole) record.secondaryRoles = p.secondaryRole;
        if (p.profilePicture) record.profilePicture = p.profilePicture;
        if (p.eod) record.eod = p.eod;
        if (p.personnelContractStatus) record.personnelContractStatus = p.personnelContractStatus;
        if (p.primaryRoleStartDate) record.primaryRoleStartDate = p.primaryRoleStartDate;
        if (p.departureDate) record.departureDate = p.departureDate;
        if (p.departureMeetingDate) record.departureMeetingDate = p.departureMeetingDate;
        if (p.contractLcat) record.contractLcat = p.contractLcat;
        if (p.location) record.location = p.location;
        if (p.tenure) record.tenure = p.tenure;
        if (p.portfolioLeadership) record.portfolioLeadership = p.portfolioLeadership;
        
        // For vacant positions from Team List (departed personnel)
        if (p.nodeType === 'vacant') {
          record.targetHireDate = '';
          record.requirements = '';
        }
        
        return record;
      });
      
      // 5. Add vacant positions from Vacant Positions sheet (positions that never had a person)
      const vacants = (typeof TeamsCrudService !== 'undefined') ? TeamsCrudService.getAllVacantPositions() : [];
      const vacantRecords = vacants.map(v => {
        // Determine parent: use Supervisor UPID if available, else fall back to team/task
        let vParentId = v.supervisorUpid || null;
        if (!vParentId || !upidMap.has(vParentId)) {
          if (v.team) {
            vParentId = ID_PREFIX_TEAM + v.team;
          } else if (v.task) {
            vParentId = ID_PREFIX_TASK + v.task;
          } else {
            vParentId = ID_ROOT;
          }
        }
        
        return {
          id: v.vacantId,
          parentId: vParentId,
          company: '',
          contract: v.contract,
          task: v.task || null,
          team: v.team || null,
          name: 'VACANT',
          email: '',
          title: v.title || 'Vacant Position',
          type: 'vacant',
          active: true,
          targetHireDate: v.targetHireDate || '',
          requirements: v.requirements || '',
          _isVacantPosition: true
        };
      });
      
      // 6. Combine all nodes
      const allRecords = [rootNode, ...taskNodes, ...teamNodes, ...personRecords, ...vacantRecords];
      
      // Cache the result
      if (typeof AppCache !== 'undefined') {
        AppCache.set(CACHE_KEY_ALL_DATA, allRecords, CACHE_TTL);
        console.log('[OrgCrudService] Org data cached:', allRecords.length, 'total nodes');
      }
      
      return allRecords;
      
    } catch (error) {
      console.error('[OrgCrudService] Error in getAllData:', error);
      return [];
    }
  }
  
  /**
   * Get org chart data as JSON string
   * @returns {string} JSON string of org chart data
   */
  function getAllDataJson() {
    return JSON.stringify(getAllData());
  }
  
  /**
   * Get a node by ID (UPID for people, synthetic for structural)
   * @param {string} id - Node ID
   * @returns {Object|null} Node object or null
   */
  function getById(id) {
    const data = getAllData();
    return data.find(p => p.id === String(id)) || null;
  }
  
  // ============================================================================
  // PUBLIC: Write Operations (to external Team List)
  // ============================================================================
  
  /**
   * Build a row array from a person object matching Team List column order
   * @private
   * @param {Object} person - Person data object
   * @returns {Array} Row array for Team List sheet
   */
  function _buildTeamListRow(person) {
    return [
      person.employeeCode || '',
      person.company || '',
      person.contract || '',
      person.task || '',
      person.primaryWorkstream || person.team || '',
      person.secondaryWorkstream || '',
      person.firstName || '',
      person.lastName || '',
      person.email || '',
      person.primaryRole || '',
      person.secondaryRole || person.secondaryRoles || '',
      person.primaryRoleStartDate || '',
      person.cpc || '',
      person.hid || '',
      person.upid || '',
      person.supervisorEmail || '',
      person.supervisorUpid || '',
      person.portfolioLeadership ? 'TRUE' : '',
      person.profilePicture || '',
      person.eod || '',
      person.personnelContractStatus || 'Active',
      person.departureDate || '',
      person.departureMeetingDate || '',
      person.contractLcat || '',
      person.location || '',
      person.tenure || '',
      person.nodeType || person.type || 'person',
      person.activeInOrg !== false ? 'TRUE' : 'FALSE'
    ];
  }
  
  /**
   * Add a new person to the Team List
   * @param {Object} person - Person data
   * @returns {Object} Result with success status
   */
  function addPerson(person) {
    PermissionService.requirePermission('org.edit');
    
    try {
      const sheet = _getTeamListSheet();
      
      // Generate UPID if not provided
      if (!person.upid) {
        const cpc = person.cpc || '400';
        const hid = person.hid || _generateNextHid();
        person.upid = cpc + '-' + hid;
      }
      
      // Split name into first/last if provided as single field
      if (person.name && (!person.firstName || !person.lastName)) {
        const parts = person.name.split(' ');
        person.firstName = person.firstName || parts[0] || '';
        person.lastName = person.lastName || parts.slice(1).join(' ') || '';
      }
      
      // Map 'title' → 'primaryRole' for frontend compatibility
      if (person.title && !person.primaryRole) {
        person.primaryRole = person.title;
      }
      
      // Map 'reportsTo' → 'supervisorUpid' for frontend compatibility
      if (person.reportsTo && !person.supervisorUpid) {
        person.supervisorUpid = person.reportsTo;
      }
      
      const newRow = _buildTeamListRow(person);
      sheet.appendRow(newRow);
      
      _invalidateCache();
      
      console.log('[OrgCrudService] Added person:', person.upid);
      return { success: true, id: person.upid };
    } catch (error) {
      console.error('[OrgCrudService] Error adding person:', error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Generate next available HID
   * @private
   * @returns {string} 3-digit HID string
   */
  function _generateNextHid() {
    const personnel = _readTeamListRaw();
    let maxHid = 0;
    personnel.forEach(p => {
      const num = parseInt(p.hid, 10);
      if (!isNaN(num) && num > maxHid) maxHid = num;
    });
    return String(maxHid + 1).padStart(3, '0');
  }
  
  /**
   * Update an existing person in the Team List
   * @param {string} id - UPID of the person
   * @param {Object} updates - Fields to update
   * @returns {Object} Result with success status
   */
  function updatePerson(id, updates) {
    PermissionService.requirePermission('org.edit');
    
    try {
      const sheet = _getTeamListSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Build column map
      const colMap = {};
      headers.forEach((h, i) => colMap[String(h).trim()] = i);
      
      // Find the row by UPID
      const upidCol = colMap['Unique Personnel ID (UPID)'];
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][upidCol]).trim() === String(id).trim()) {
          rowIndex = i + 1; // 1-indexed
          break;
        }
      }
      
      if (rowIndex === -1) {
        return { success: false, error: 'Person not found with UPID: ' + id };
      }
      
      // Map update field names to Team List header names
      const fieldToHeader = {
        employeeCode: 'Employee Code',
        company: 'Company',
        contract: 'Contract',
        task: 'Task',
        primaryWorkstream: 'Primary Workstream',
        team: 'Primary Workstream',
        secondaryWorkstream: 'Secondary Workstream',
        firstName: 'First Name',
        lastName: 'Last Name',
        email: 'Email',
        title: 'Primary Role',
        primaryRole: 'Primary Role',
        secondaryRole: 'Secondary Role',
        secondaryRoles: 'Secondary Role',
        primaryRoleStartDate: 'Primary Role Start Date',
        cpc: 'Contract Personnel Code (CPC)',
        hid: 'Heirarchy Identifier (HID)',
        upid: 'Unique Personnel ID (UPID)',
        supervisorEmail: 'Supervisor Email',
        supervisorUpid: 'Supervisor UPID',
        reportsTo: 'Supervisor UPID',
        portfolioLeadership: 'Portfolio Leadership?',
        profilePicture: 'Profile Picture',
        eod: 'EOD',
        personnelContractStatus: 'Personnel Contract Status',
        departureDate: 'Departure Date',
        departureMeetingDate: 'Departure Meeting Date',
        contractLcat: 'Contract LCAT',
        location: 'Location (City, ST)',
        nodeType: 'Node Type',
        type: 'Node Type',
        activeInOrg: 'Active In Org'
      };
      
      // Handle 'name' field → split into First Name / Last Name
      if (updates.name !== undefined) {
        const parts = updates.name.split(' ');
        updates.firstName = parts[0] || '';
        updates.lastName = parts.slice(1).join(' ') || '';
      }
      
      // Handle 'parentId' → Supervisor UPID
      if (updates.parentId !== undefined) {
        // Only set supervisorUpid if it looks like a UPID (not a structural ID)
        if (!updates.parentId.startsWith('task:') && !updates.parentId.startsWith('team:') && updates.parentId !== 'root') {
          updates.supervisorUpid = updates.parentId;
        }
      }
      
      // Update each field
      Object.keys(updates).forEach(field => {
        const headerName = fieldToHeader[field];
        if (headerName && colMap[headerName] !== undefined) {
          const col = colMap[headerName] + 1; // 1-indexed
          let value = updates[field];
          if (value === null || value === undefined) value = '';
          sheet.getRange(rowIndex, col).setValue(value);
        }
      });
      
      _invalidateCache();
      
      console.log('[OrgCrudService] Updated person:', id);
      return { success: true, id: id };
    } catch (error) {
      console.error('[OrgCrudService] Error updating person:', error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Delete a person from the Team List
   * Performs a soft delete by setting Active In Org = FALSE and Personnel Contract Status = 'Departed'
   * @param {string} id - UPID of the person (or vacant ID)
   * @returns {Object} Result with success status
   */
  function deletePerson(id) {
    PermissionService.requirePermission('org.edit');
    
    try {
      // Check if this is a vacant position from the Vacant Positions sheet
      if (id && id.startsWith('VAC-')) {
        if (typeof TeamsCrudService !== 'undefined') {
          return TeamsCrudService.deleteVacantPosition(id);
        }
        return { success: false, error: 'Cannot delete vacant position: TeamsCrudService not available' };
      }
      
      // Otherwise it's a person in the Team List — soft delete
      const sheet = _getTeamListSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      const colMap = {};
      headers.forEach((h, i) => colMap[String(h).trim()] = i);
      
      const upidCol = colMap['Unique Personnel ID (UPID)'];
      const activeInOrgCol = colMap['Active In Org'];
      const statusCol = colMap['Personnel Contract Status'];
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][upidCol]).trim() === String(id).trim()) {
          const rowIndex = i + 1;
          // Soft delete: mark as departed and inactive in org
          sheet.getRange(rowIndex, activeInOrgCol + 1).setValue('FALSE');
          sheet.getRange(rowIndex, statusCol + 1).setValue('Departed');
          
          _invalidateCache();
          
          console.log('[OrgCrudService] Soft-deleted person:', id);
          return { success: true, id: id };
        }
      }
      
      return { success: false, error: 'Person not found with UPID: ' + id };
    } catch (error) {
      console.error('[OrgCrudService] Error deleting person:', error);
      return { success: false, error: error.message };
    }
  }
  
  // ============================================================================
  // PUBLIC: Query Operations
  // ============================================================================
  
  /**
   * Get unique task IDs for filtering
   * @returns {Array<string>} Array of unique task IDs
   */
  function getTasks() {
    const data = getAllData();
    const tasks = new Set();
    data.forEach(node => {
      if (node.task) tasks.add(node.task);
    });
    return Array.from(tasks).sort();
  }
  
  /**
   * Get node type configuration
   * @returns {Object} Node type config with labels
   */
  function getNodeTypeConfig() {
    return {
      types: ['hidden', 'director', 'deputy', 'task', 'team', 'lead', 'person', 'vacant'],
      labels: {
        hidden: 'Hidden Root',
        director: 'Director',
        deputy: 'Deputy',
        task: 'Task',
        team: 'Team',
        lead: 'Team Lead',
        person: 'Person',
        vacant: 'Vacant'
      }
    };
  }
  
  /**
   * Get task color configuration
   * Reads from Team Mappings metadata
   * @returns {Object} Task ID to color mapping
   */
  function getTaskColors() {
    if (typeof TeamsCrudService !== 'undefined') {
      return TeamsCrudService.getTaskColors();
    }
    return { 'default': '#95a5a6' };
  }
  
  /**
   * Gets task friendly names from Team Mappings
   * @returns {Object} Map of task IDs to display names
   */
  function getTaskFriendlyNames() {
    try {
      if (typeof TeamsCrudService !== 'undefined') {
        return TeamsCrudService.getTaskFriendlyNames();
      }
      return {};
    } catch (error) {
      console.error('[OrgCrudService] Error getting task friendly names:', error);
      return {};
    }
  }
  
  /**
   * Gets management emails derived from personnel data
   * Finds directors, deputies, task leads, and team leads by Node Type
   * @returns {Object} Map of task IDs to management email objects
   */
  function getManagementEmails() {
    try {
      const nodes = getAllData();
      const managementEmails = {};
      
      // Find directors and deputies by contract
      const contractManagers = {};
      nodes.forEach(node => {
        if (node.type === 'director' && node.contract && node.email) {
          if (!contractManagers[node.contract]) contractManagers[node.contract] = {};
          contractManagers[node.contract].director = node.email;
        }
        if (node.type === 'deputy' && node.contract && node.email) {
          if (!contractManagers[node.contract]) contractManagers[node.contract] = {};
          contractManagers[node.contract].deputy = node.email;
        }
      });
      
      // Build management email map per task
      const taskIds = getTasks();
      taskIds.forEach(taskId => {
        // Find the task node to get its contract
        const taskNode = nodes.find(n => n.type === 'task' && n.task === taskId);
        const contract = taskNode ? taskNode.contract : '';
        
        managementEmails[taskId] = {
          contractManagerEmail: contractManagers[contract]?.director || '',
          deputyManagerEmail: contractManagers[contract]?.deputy || '',
          taskLeadEmail: '',
          teamLeadEmail: ''
        };
      });
      
      // Find leads per task from personnel
      nodes.forEach(node => {
        if (!node.task || !node.email) return;
        
        // Task lead: look for the node whose parentId points to the task structural node
        // and who has a lead-like role
        if (node.type === 'lead' && managementEmails[node.task]) {
          if (!managementEmails[node.task].teamLeadEmail) {
            managementEmails[node.task].teamLeadEmail = node.email;
          }
        }
      });
      
      return managementEmails;
    } catch (error) {
      console.error('[OrgCrudService] Error getting management emails:', error);
      return {};
    }
  }
  
  /**
   * Get module summary for dashboard
   * @returns {Object} Summary statistics
   */
  function getModuleSummary() {
    const data = getAllData();
    
    // Exclude structural nodes from counts
    const nonStructural = data.filter(n => !n._isStructural);
    
    const stats = {
      total: nonStructural.length,
      byType: {},
      vacantCount: 0,
      taskCount: 0
    };
    
    nonStructural.forEach(node => {
      stats.byType[node.type] = (stats.byType[node.type] || 0) + 1;
      if (node.type === 'vacant') stats.vacantCount++;
    });
    
    stats.taskCount = getTasks().length;
    
    return stats;
  }
  
  /**
   * Get total count of people (for dashboard stats)
   * @returns {number}
   */
  function getTotalCount() {
    try {
      const data = getAllData();
      return data.filter(n => !n._isStructural && n.type !== 'hidden').length;
    } catch (error) {
      console.error('[OrgCrudService] Error getting total count:', error);
      return 0;
    }
  }
  
  /**
   * Get person(s) by email address
   * @param {string} email - Email address to look up
   * @returns {Array<Object>} Array of person objects, empty if not found
   */
  function getByEmail(email) {
    if (!email) return [];
    const normalizedEmail = email.toLowerCase().trim();
    const data = getAllData();
    return data.filter(p => p.email && p.email.toLowerCase().trim() === normalizedEmail);
  }
  
  /**
   * Get detailed user profile from org chart
   * Aggregates data if user appears multiple times (multiple tasks/teams)
   * @param {string} email - Email address to look up
   * @returns {Object|null} User profile with aggregated details, or null if not found
   */
  function getUserProfile(email) {
    const entries = getByEmail(email);
    if (entries.length === 0) return null;
    
    const primary = entries[0];
    
    const tasks = [...new Set(entries.map(e => e.task).filter(Boolean))];
    const teams = [...new Set(entries.map(e => e.team).filter(Boolean))];
    const contracts = [...new Set(entries.map(e => e.contract).filter(Boolean))];
    
    return {
      id: primary.id,
      name: primary.name,
      email: primary.email,
      title: primary.title,
      type: primary.type,
      company: primary.company,
      profilePicture: primary.profilePicture || null,
      upid: primary.upid || null,
      tasks: tasks,
      teams: teams,
      contracts: contracts,
      primaryTask: tasks[0] || null,
      primaryTeam: teams[0] || null,
      entries: entries
    };
  }
  
  // Public API
  return {
    // Constants (backward compatible)
    SHEET_NAME: TEAM_LIST_SHEET_NAME,
    HEADERS: HEADERS,
    TEAM_LIST_HEADERS: TEAM_LIST_HEADERS,
    // Core reads
    getAllData: getAllData,
    getAllDataJson: getAllDataJson,
    getById: getById,
    getByEmail: getByEmail,
    getUserProfile: getUserProfile,
    // Writes
    addPerson: addPerson,
    updatePerson: updatePerson,
    deletePerson: deletePerson,
    // Queries
    getTasks: getTasks,
    getNodeTypeConfig: getNodeTypeConfig,
    getTaskColors: getTaskColors,
    getTaskFriendlyNames: getTaskFriendlyNames,
    getManagementEmails: getManagementEmails,
    getModuleSummary: getModuleSummary,
    getTotalCount: getTotalCount
  };
})();
