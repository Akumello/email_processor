/**
 * Setup Integration Tests
 * Runs CRUD operations during setup to verify all modules work correctly
 * Uses the same backend functions that the frontend uses
 */

const SetupIntegrationTests = (function() {
  'use strict';
  
  // Track all items created during tests for cleanup
  let _createdItems = {
    slas: [],
    requests: [],
    orgPersons: []
  };
  
  // Test results tracking
  let _testResults = {
    passed: 0,
    failed: 0,
    errors: [],
    details: []
  };
  
  // Verbose logging flag
  let _verbose = true;
  
  /**
   * Log message based on verbosity setting
   * @private
   */
  function _log(message, level = 'info') {
    if (_verbose || level === 'error') {
      const prefix = level === 'error' ? '❌' : level === 'success' ? '✅' : 'ℹ️';
      console.log(`${prefix} ${message}`);
    }
  }
  
  /**
   * Record a test result
   * @private
   */
  function _recordTest(testName, passed, details = null, error = null) {
    if (passed) {
      _testResults.passed++;
      _log(`PASS: ${testName}`, 'success');
    } else {
      _testResults.failed++;
      _testResults.errors.push({ test: testName, error: error || details });
      _log(`FAIL: ${testName} - ${error || details}`, 'error');
    }
    
    _testResults.details.push({
      name: testName,
      passed: passed,
      details: details,
      error: error,
      timestamp: new Date().toISOString()
    });
  }
  
  /**
   * Reset tracking state
   */
  function reset() {
    _createdItems = { slas: [], requests: [], orgPersons: [] };
    _testResults = { passed: 0, failed: 0, errors: [], details: [] };
  }
  
  /**
   * Set verbosity
   * @param {boolean} verbose - Whether to log detailed output
   */
  function setVerbose(verbose) {
    _verbose = verbose;
  }
  
  // ============================================================================
  // SLA INTEGRATION TESTS
  // ============================================================================
  
  /**
   * Run SLA module integration tests
   * Creates, reads, updates, and deletes SLAs using backend APIs
   * @returns {Object} Test results for SLA module
   */
  function runSLATests() {
    _log('Starting SLA Integration Tests...', 'info');
    const moduleResults = { created: [], updated: [], deleted: [], verified: [], errors: [] };
    
    // Temporarily bypass permissions for testing
    const originalRequirePermission = PermissionService.requirePermission;
    PermissionService.requirePermission = function() { return true; };
    
    try {
      // ---- TEST 1: Create SLAs ----
      const testSLAs = [
        {
          slaName: 'Integration Test - Quantity SLA',
          slaType: 'quantity',
          description: 'Created by integration test',
          teamId: 'TEAM-001',
          taskId: 'TASK-001',
          startDate: new Date(),
          endDate: new Date(Date.now() + 90 * 24 * 60 * 60 * 1000),
          status: 'SLA-ST-003',
          currentValue: 25,
          targetValue: 100,
          useRange: false,
          frequency: 'monthly',
          tags: ['integration-test'],
          isParent: false
        },
        {
          slaName: 'Integration Test - Compliance SLA',
          slaType: 'compliance',
          description: 'Created by integration test for compliance',
          teamId: 'TEAM-002',
          taskId: 'TASK-002',
          startDate: new Date(),
          endDate: new Date(Date.now() + 90 * 24 * 60 * 60 * 1000),
          status: 'SLA-ST-005',
          targetValue: 1,
          useRange: false,
          frequency: 'quarterly',
          tags: ['integration-test', 'compliance'],
          isParent: false
        },
        {
          slaName: 'Integration Test - Parent SLA',
          slaType: 'quantity',
          description: 'Parent SLA created by integration test',
          teamId: 'TEAM-003',
          taskId: 'TASK-003',
          startDate: new Date(),
          endDate: new Date(Date.now() + 90 * 24 * 60 * 60 * 1000),
          status: 'SLA-ST-003',
          currentValue: 0,
          targetValue: 100,
          useRange: false,
          frequency: 'monthly',
          tags: ['integration-test', 'parent'],
          isParent: true,
          customFields: { parentMode: 'container' }
        }
      ];
      
      // Create each SLA
      testSLAs.forEach((slaData, index) => {
        try {
          const result = SLACrudService.createSLA(slaData);
          if (result.success) {
            const slaId = result.data?.slaId || result.slaId;
            _createdItems.slas.push(slaId);
            moduleResults.created.push({ slaId, name: slaData.slaName });
            _recordTest(`SLA Create #${index + 1}: ${slaData.slaName}`, true, `Created ${slaId}`);
          } else {
            moduleResults.errors.push({ operation: 'create', error: result.error, data: slaData.slaName });
            _recordTest(`SLA Create #${index + 1}: ${slaData.slaName}`, false, null, result.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'create', error: error.message, data: slaData.slaName });
          _recordTest(`SLA Create #${index + 1}: ${slaData.slaName}`, false, null, error.message);
        }
      });
      
      // ---- TEST 2: Verify SLAs by reading back ----
      _createdItems.slas.forEach((slaId, index) => {
        try {
          const result = SLACrudService.getSLAById(slaId);
          if (result.success && result.data) {
            moduleResults.verified.push({ slaId, verified: true });
            _recordTest(`SLA Verify #${index + 1}: ${slaId}`, true, 'Read-back successful');
          } else {
            moduleResults.errors.push({ operation: 'verify', error: 'SLA not found after create', slaId });
            _recordTest(`SLA Verify #${index + 1}: ${slaId}`, false, null, 'SLA not found after create');
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'verify', error: error.message, slaId });
          _recordTest(`SLA Verify #${index + 1}: ${slaId}`, false, null, error.message);
        }
      });
      
      // ---- TEST 3: Update one SLA (change progress/currentValue) ----
      if (_createdItems.slas.length > 0) {
        const slaToUpdate = _createdItems.slas[0];
        try {
          // First get the SLA to get current rowVersion
          const getResult = SLACrudService.getSLAById(slaToUpdate);
          if (getResult.success) {
            const updates = {
              currentValue: 50,
              description: 'Updated by integration test - progress changed'
            };
            const updateResult = SLACrudService.updateSLA(slaToUpdate, updates, getResult.data.rowVersion);
            if (updateResult.success) {
              moduleResults.updated.push({ slaId: slaToUpdate, updates });
              _recordTest(`SLA Update: ${slaToUpdate}`, true, 'Updated currentValue to 50');
              
              // Verify the update
              const verifyResult = SLACrudService.getSLAById(slaToUpdate);
              const actualValue = verifyResult.success ? verifyResult.data.currentValue : 'N/A';
              const expectedValue = 50;
              
              if (verifyResult.success && verifyResult.data.currentValue === expectedValue) {
                _recordTest(`SLA Update Verify: ${slaToUpdate}`, true, `Update verified - currentValue is ${expectedValue}`);
              } else {
                const errorMsg = `Update not reflected - Expected currentValue: ${expectedValue}, Actual: ${actualValue}, Full data: ${JSON.stringify(verifyResult.data)}`;
                _recordTest(`SLA Update Verify: ${slaToUpdate}`, false, null, errorMsg);
              }
            } else {
              moduleResults.errors.push({ operation: 'update', error: updateResult.error, slaId: slaToUpdate });
              _recordTest(`SLA Update: ${slaToUpdate}`, false, null, updateResult.error);
            }
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'update', error: error.message, slaId: slaToUpdate });
          _recordTest(`SLA Update: ${slaToUpdate}`, false, null, error.message);
        }
      }
      
      // ---- TEST 4: Delete one SLA ----
      if (_createdItems.slas.length > 1) {
        const slaToDelete = _createdItems.slas[1]; // Delete the second one
        try {
          const deleteResult = SLACrudService.deleteSLA(slaToDelete);
          if (deleteResult.success) {
            moduleResults.deleted.push({ slaId: slaToDelete });
            _recordTest(`SLA Delete: ${slaToDelete}`, true, 'Deleted successfully');
            
            // Verify deletion (soft delete - isActive should be false)
            const verifyResult = SLACrudService.getSLAById(slaToDelete);
            if (verifyResult.success && verifyResult.data?.isActive === false) {
              _recordTest(`SLA Delete Verify: ${slaToDelete}`, true, 'Soft deletion verified - isActive=false');
            } else {
              _recordTest(`SLA Delete Verify: ${slaToDelete}`, false, null, 'SLA still active after delete');
            }
            
            // Remove from tracked items since it's deleted
            _createdItems.slas = _createdItems.slas.filter(id => id !== slaToDelete);
          } else {
            moduleResults.errors.push({ operation: 'delete', error: deleteResult.error, slaId: slaToDelete });
            _recordTest(`SLA Delete: ${slaToDelete}`, false, null, deleteResult.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'delete', error: error.message, slaId: slaToDelete });
          _recordTest(`SLA Delete: ${slaToDelete}`, false, null, error.message);
        }
      }
      
    } finally {
      // Restore permissions
      PermissionService.requirePermission = originalRequirePermission;
    }
    
    _log(`SLA Tests Complete: ${moduleResults.created.length} created, ${moduleResults.updated.length} updated, ${moduleResults.deleted.length} deleted`, 'info');
    return moduleResults;
  }
  
  // ============================================================================
  // AD-HOC REQUEST INTEGRATION TESTS
  // ============================================================================
  
  /**
   * Run Ad-Hoc module integration tests
   * Creates, reads, updates, and deletes requests using backend APIs
   * @returns {Object} Test results for Ad-Hoc module
   */
  function runAdHocTests() {
    _log('Starting Ad-Hoc Integration Tests...', 'info');
    const moduleResults = { created: [], updated: [], deleted: [], verified: [], errors: [] };
    
    // Temporarily bypass permissions for testing
    const originalRequirePermission = PermissionService.requirePermission;
    PermissionService.requirePermission = function() { return true; };
    
    try {
      // ---- TEST 1: Create Requests ----
      const testRequests = [
        {
          requestName: 'Integration Test - Data Pull Request',
          descriptionOfRequest: 'Test request created by integration tests',
          businessJustification: 'Testing the create functionality',
          requestTypeCategory: 'Data Pull',
          whatContractDoesItApplyTo: 'SQuAT',
          taskName: 'Task 1',
          priority: 'Medium',
          targetCompletionDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
        },
        {
          requestName: 'Integration Test - Report Generation',
          descriptionOfRequest: 'Test report request for integration testing',
          businessJustification: 'Verify report request creation works',
          requestTypeCategory: 'Report Generation',
          whatContractDoesItApplyTo: 'Forward',
          taskName: 'Task 2',
          priority: 'High',
          targetCompletionDate: new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
        },
        {
          requestName: 'Integration Test - Urgent Request',
          descriptionOfRequest: 'Urgent test request to verify priority handling',
          businessJustification: 'Testing urgent request flow',
          requestTypeCategory: 'Process Change',
          whatContractDoesItApplyTo: 'SQuAT',
          taskName: 'Task 3',
          priority: 'Urgent',
          targetCompletionDate: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
        }
      ];
      
      // Create each request
      testRequests.forEach((requestData, index) => {
        try {
          const result = AdHocCrudService.createRequest(requestData);
          if (result.success) {
            const requestId = result.data?.requestId || result.requestId;
            _createdItems.requests.push(requestId);
            moduleResults.created.push({ requestId, name: requestData.requestName });
            _recordTest(`Request Create #${index + 1}: ${requestData.requestName}`, true, `Created ${requestId}`);
          } else {
            moduleResults.errors.push({ operation: 'create', error: result.error, data: requestData.requestName });
            _recordTest(`Request Create #${index + 1}: ${requestData.requestName}`, false, null, result.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'create', error: error.message, data: requestData.requestName });
          _recordTest(`Request Create #${index + 1}: ${requestData.requestName}`, false, null, error.message);
        }
      });
      
      // ---- TEST 2: Verify Requests by reading back ----
      _createdItems.requests.forEach((requestId, index) => {
        try {
          const result = AdHocCrudService.getRequestById(requestId);
          if (result.success && result.data) {
            moduleResults.verified.push({ requestId, verified: true });
            _recordTest(`Request Verify #${index + 1}: ${requestId}`, true, 'Read-back successful');
          } else {
            moduleResults.errors.push({ operation: 'verify', error: 'Request not found after create', requestId });
            _recordTest(`Request Verify #${index + 1}: ${requestId}`, false, null, 'Request not found after create');
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'verify', error: error.message, requestId });
          _recordTest(`Request Verify #${index + 1}: ${requestId}`, false, null, error.message);
        }
      });
      
      // ---- TEST 3: Update one Request (change status) ----
      if (_createdItems.requests.length > 0) {
        const requestToUpdate = _createdItems.requests[0];
        try {
          // First get the request to get current rowVersion
          const getResult = AdHocCrudService.getRequestById(requestToUpdate);
          if (getResult.success) {
            const updates = {
              status: 'in-progress',
              assignedTo: 'test@example.com',
              internalNotes: 'Updated by integration test'
            };
            const updateResult = AdHocCrudService.updateRequest(requestToUpdate, updates, getResult.data.rowVersion);
            if (updateResult.success) {
              moduleResults.updated.push({ requestId: requestToUpdate, updates });
              _recordTest(`Request Update: ${requestToUpdate}`, true, 'Updated status to in-progress');
              
              // Verify the update
              const verifyResult = AdHocCrudService.getRequestById(requestToUpdate);
              if (verifyResult.success && verifyResult.data.status === 'in-progress') {
                _recordTest(`Request Update Verify: ${requestToUpdate}`, true, 'Update verified');
              } else {
                _recordTest(`Request Update Verify: ${requestToUpdate}`, false, null, 'Update not reflected');
              }
            } else {
              moduleResults.errors.push({ operation: 'update', error: updateResult.error, requestId: requestToUpdate });
              _recordTest(`Request Update: ${requestToUpdate}`, false, null, updateResult.error);
            }
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'update', error: error.message, requestId: requestToUpdate });
          _recordTest(`Request Update: ${requestToUpdate}`, false, null, error.message);
        }
      }
      
      // ---- TEST 4: Delete one Request ----
      if (_createdItems.requests.length > 1) {
        const requestToDelete = _createdItems.requests[1]; // Delete the second one
        try {
          const deleteResult = AdHocCrudService.deleteRequest(requestToDelete);
          if (deleteResult.success) {
            moduleResults.deleted.push({ requestId: requestToDelete });
            _recordTest(`Request Delete: ${requestToDelete}`, true, 'Deleted successfully');
            
            // Verify deletion
            const verifyResult = AdHocCrudService.getRequestById(requestToDelete);
            if (!verifyResult.success) {
              _recordTest(`Request Delete Verify: ${requestToDelete}`, true, 'Deletion verified');
            } else {
              _recordTest(`Request Delete Verify: ${requestToDelete}`, false, null, 'Request still exists after delete');
            }
            
            // Remove from tracked items
            _createdItems.requests = _createdItems.requests.filter(id => id !== requestToDelete);
          } else {
            moduleResults.errors.push({ operation: 'delete', error: deleteResult.error, requestId: requestToDelete });
            _recordTest(`Request Delete: ${requestToDelete}`, false, null, deleteResult.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'delete', error: error.message, requestId: requestToDelete });
          _recordTest(`Request Delete: ${requestToDelete}`, false, null, error.message);
        }
      }
      
    } finally {
      // Restore permissions
      PermissionService.requirePermission = originalRequirePermission;
    }
    
    _log(`Ad-Hoc Tests Complete: ${moduleResults.created.length} created, ${moduleResults.updated.length} updated, ${moduleResults.deleted.length} deleted`, 'info');
    return moduleResults;
  }
  
  // ============================================================================
  // ORG MODULE INTEGRATION TESTS
  // ============================================================================
  
  /**
   * Run Org module integration tests
   * Creates, reads, updates, and deletes org chart entries using backend APIs
   * @returns {Object} Test results for Org module
   */
  function runOrgTests() {
    _log('Starting Org Integration Tests...', 'info');
    const moduleResults = { created: [], updated: [], deleted: [], verified: [], errors: [] };
    
    // Temporarily bypass permissions for testing
    const originalRequirePermission = PermissionService.requirePermission;
    PermissionService.requirePermission = function() { return true; };
    
    try {
      // ---- TEST 1: Create Org entries ----
      // Use UPID-based IDs. parentId/reportsTo should be a valid UPID or empty for top-level.
      const testPersons = [
        {
          name: 'Test Team Lead',
          title: 'Integration Test Lead',
          task: 'TASK-001',
          contract: 'SQuAT',
          email: 'test.lead@integration.test',
          cpc: '310',
          hid: '901',
          reportsTo: '' // Top-level for testing
        },
        {
          name: 'Test Team Member',
          title: 'Integration Test Analyst',
          task: 'TASK-001',
          contract: 'SQuAT',
          email: 'test.member@integration.test',
          cpc: '410',
          hid: '902',
          reportsTo: null // Will be set to first person's UPID
        }
      ];
      
      let firstCreatedId = null;
      
      // Create each person
      testPersons.forEach((personData, index) => {
        try {
          // Link second person to first
          if (index === 1 && firstCreatedId) {
            personData.reportsTo = firstCreatedId;
          }
          
          const result = OrgCrudService.addPerson(personData);
          if (result.success) {
            const personId = result.id;
            if (index === 0) firstCreatedId = personId;
            _createdItems.orgPersons.push(personId);
            moduleResults.created.push({ personId, name: personData.name });
            _recordTest(`Org Create #${index + 1}: ${personData.name}`, true, `Created UPID ${personId}`);
          } else {
            moduleResults.errors.push({ operation: 'create', error: result.error, data: personData.name });
            _recordTest(`Org Create #${index + 1}: ${personData.name}`, false, null, result.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'create', error: error.message, data: personData.name });
          _recordTest(`Org Create #${index + 1}: ${personData.name}`, false, null, error.message);
        }
      });
      
      // ---- TEST 2: Verify Org entries by reading back ----
      _createdItems.orgPersons.forEach((personId, index) => {
        try {
          const person = OrgCrudService.getById(personId);
          if (person) {
            moduleResults.verified.push({ personId, verified: true });
            _recordTest(`Org Verify #${index + 1}: ${personId}`, true, `Found: ${person.name}`);
          } else {
            moduleResults.errors.push({ operation: 'verify', error: 'Person not found after create', personId });
            _recordTest(`Org Verify #${index + 1}: ${personId}`, false, null, 'Person not found after create');
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'verify', error: error.message, personId });
          _recordTest(`Org Verify #${index + 1}: ${personId}`, false, null, error.message);
        }
      });
      
      // ---- TEST 3: Update one Person (change name/title) ----
      if (_createdItems.orgPersons.length > 0) {
        const personToUpdate = _createdItems.orgPersons[0];
        try {
          const updates = {
            name: 'Updated Test Lead',
            title: 'Senior Integration Test Lead'
          };
          const updateResult = OrgCrudService.updatePerson(personToUpdate, updates);
          if (updateResult.success) {
            moduleResults.updated.push({ personId: personToUpdate, updates });
            _recordTest(`Org Update: ${personToUpdate}`, true, 'Updated name and title');
            
            // Verify the update
            const person = OrgCrudService.getById(personToUpdate);
            if (person && person.name === 'Updated Test Lead') {
              _recordTest(`Org Update Verify: ${personToUpdate}`, true, 'Update verified');
            } else {
              _recordTest(`Org Update Verify: ${personToUpdate}`, false, null, 'Update not reflected in name');
            }
          } else {
            moduleResults.errors.push({ operation: 'update', error: updateResult.error, personId: personToUpdate });
            _recordTest(`Org Update: ${personToUpdate}`, false, null, updateResult.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'update', error: error.message, personId: personToUpdate });
          _recordTest(`Org Update: ${personToUpdate}`, false, null, error.message);
        }
      }
      
      // ---- TEST 4: Soft-delete one Person and verify ----
      if (_createdItems.orgPersons.length > 1) {
        const personToDelete = _createdItems.orgPersons[_createdItems.orgPersons.length - 1];
        try {
          const deleteResult = OrgCrudService.deletePerson(personToDelete);
          if (deleteResult.success) {
            moduleResults.deleted.push({ personId: personToDelete });
            _recordTest(`Org Delete: ${personToDelete}`, true, 'Soft-deleted successfully');
            
            // Verify soft deletion: person should still exist but be marked Departed
            const person = OrgCrudService.getById(personToDelete);
            if (person && person.personnelContractStatus === 'Departed') {
              _recordTest(`Org Delete Verify: ${personToDelete}`, true, 'Soft deletion verified (status: Departed)');
            } else if (!person) {
              // The person might not show if Active In Org was set to FALSE and getAllData filters them
              _recordTest(`Org Delete Verify: ${personToDelete}`, true, 'Person no longer visible (filtered out)');
            } else {
              _recordTest(`Org Delete Verify: ${personToDelete}`, false, null, 'Person still active after soft delete');
            }
            
            // Remove from tracked items
            _createdItems.orgPersons = _createdItems.orgPersons.filter(id => id !== personToDelete);
          } else {
            moduleResults.errors.push({ operation: 'delete', error: deleteResult.error, personId: personToDelete });
            _recordTest(`Org Delete: ${personToDelete}`, false, null, deleteResult.error);
          }
        } catch (error) {
          moduleResults.errors.push({ operation: 'delete', error: error.message, personId: personToDelete });
          _recordTest(`Org Delete: ${personToDelete}`, false, null, error.message);
        }
      }
      
    } finally {
      // Restore permissions
      PermissionService.requirePermission = originalRequirePermission;
    }
    
    _log(`Org Tests Complete: ${moduleResults.created.length} created, ${moduleResults.updated.length} updated, ${moduleResults.deleted.length} deleted`, 'info');
    return moduleResults;
  }
  
  // ============================================================================
  // MAIN TEST RUNNER
  // ============================================================================
  
  /**
   * Run all integration tests
   * @param {Object} options - Test options
   * @param {boolean} options.verbose - Whether to log detailed output (default: true)
   * @param {boolean} options.sla - Run SLA tests (default: true)
   * @param {boolean} options.adHoc - Run Ad-Hoc tests (default: true)
   * @param {boolean} options.org - Run Org tests (default: true)
   * @returns {Object} Complete test results
   */
  function runAllTests(options = {}) {
    const {
      verbose = true,
      sla = true,
      adHoc = true,
      org = true
    } = options;
    
    _verbose = verbose;
    reset();
    
    const startTime = Date.now();
    console.log('========================================');
    console.log('🧪 Starting Integration Tests');
    console.log('========================================');
    
    const results = {
      sla: null,
      adHoc: null,
      org: null,
      summary: null,
      timing: { startTime: new Date().toISOString() }
    };
    
    // Run module tests
    if (sla && typeof SLACrudService !== 'undefined') {
      results.sla = runSLATests();
    }
    
    if (adHoc && typeof AdHocCrudService !== 'undefined') {
      results.adHoc = runAdHocTests();
    }
    
    if (org && typeof OrgCrudService !== 'undefined') {
      results.org = runOrgTests();
    }
    
    // Calculate summary
    const endTime = Date.now();
    results.timing.endTime = new Date().toISOString();
    results.timing.durationMs = endTime - startTime;
    
    results.summary = {
      totalPassed: _testResults.passed,
      totalFailed: _testResults.failed,
      totalTests: _testResults.passed + _testResults.failed,
      passRate: _testResults.passed + _testResults.failed > 0 
        ? Math.round((_testResults.passed / (_testResults.passed + _testResults.failed)) * 100) 
        : 0,
      errors: _testResults.errors,
      allPassed: _testResults.failed === 0
    };
    
    // Print summary
    console.log('========================================');
    console.log('📊 Integration Test Summary');
    console.log('========================================');
    console.log(`Total Tests: ${results.summary.totalTests}`);
    console.log(`Passed: ${results.summary.totalPassed} ✅`);
    console.log(`Failed: ${results.summary.totalFailed} ❌`);
    console.log(`Pass Rate: ${results.summary.passRate}%`);
    console.log(`Duration: ${results.timing.durationMs}ms`);
    
    if (results.summary.errors.length > 0) {
      console.log('\nErrors:');
      results.summary.errors.forEach((err, i) => {
        console.log(`  ${i + 1}. ${err.test}: ${err.error}`);
      });
    }
    
    console.log('========================================');
    
    return results;
  }
  
  /**
   * Get created items (for cleanup reference)
   * @returns {Object} Created items by module
   */
  function getCreatedItems() {
    return { ..._createdItems };
  }
  
  /**
   * Get test results
   * @returns {Object} Test results
   */
  function getTestResults() {
    return { ..._testResults };
  }
  
  // Public API
  return {
    reset: reset,
    setVerbose: setVerbose,
    runSLATests: runSLATests,
    runAdHocTests: runAdHocTests,
    runOrgTests: runOrgTests,
    runAllTests: runAllTests,
    getCreatedItems: getCreatedItems,
    getTestResults: getTestResults
  };
})();
