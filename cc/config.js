/**
 * Configuration for Command Center Application
 */

const CONFIG = (function () {
  "use strict";

  return {
    // Spreadsheet configuration
    SPREADSHEET_ID: "YOUR_SPREADSHEET_ID_HERE",

    // Sheet names (ALL_CAPS_SNAKE_CASE convention for system-managed sheets)
    SHEETS: {
      SLA: "SLA_MASTER",
      ADHOC_REQUESTS: "ADHOC_REQUESTS",
      ADHOC_ACTIVITY_LOG: "ADHOC_ACTIVITY_LOG",
      SLA_ACTIVITY_LOG: "SLA_ACTIVITY_LOG",
      TEAMS: "TEAMS",
      LOOKUP_NOTIFICATION_EMAILS: "LOOKUP_NOTIFICATION_EMAILS",
      ORG_CHART: "ORG_CHART", // Deprecated: personnel data now in external Team List
      TEAM_LIST: "Team List", // External sheet: personnel roster (in teamlist datastore)
      TEAM_MAPPINGS: "Team Mappings", // Structure: task/team definitions and metadata
      // Notification system sheets
      NOTIFICATION_LOG: "NOTIFICATION_LOG",
      NOTIFICATION_OPTOUTS: "NOTIFICATION_OPTOUTS",
      EMAIL_TEMPLATES: "EMAIL_TEMPLATES"
    },

    // Application settings
    APP: {
      NAME: "Command Center",
      VERSION: "1.1.0",
      TIMEZONE: "America/Los_Angeles",
      DATE_FORMAT: "MMM dd, yyyy",
      DATETIME_FORMAT: "MMM dd, yyyy HH:mm",
      // Paste your deployment URL here (from Deploy > Manage deployments > Web app URL)
      DEPLOYMENT_URL: "",
      NAVIGATION: {
        MODE: "sidebar", // Options: 'top', 'sidebar'
        SIDEBAR_WIDTH: "250px",
        SIDEBAR_COLLAPSED_WIDTH: "64px",
      },
    },

    // Modules configuration
    MODULES: {
      SLA: {
        ID: 'sla',
        NAME: 'SLA Tracker',
        ENABLED: true,
        ICON: 'bi-speedometer2'
      },
      ADHOC: {
        ID: 'adhoc',
        NAME: 'Ad Hoc Requests',
        ENABLED: true,
        ICON: 'bi-envelope-paper'
      },
      ORG: {
        ID: 'org',
        NAME: 'Org Chart',
        ENABLED: true,
        ICON: 'bi-diagram-3'
      }
    },

    // SLA configuration
    SLA: {
      TYPES: [
        "quantity",
        "compliance",
        "parent",
      ],
      STATUSES: [
        "exceeded",
        "met",
        "on-track",
        "no-requests",
        "not-started",
        "at-risk",
        "not-met",
      ],
      PARENT_MODES: ["container", "tracked"],
      DEFAULT_STATUS: "not-started",
    },

    // Ad Hoc Requests configuration
    ADHOC: {
      STATUSES: [
        "New", "Assigned", "In Progress", "Pending Clarification",
        "On Hold", "In Review", "Completed", "Resolved", "Closed",
        "Cancelled", "Rejected"
      ],
      PRIORITIES: ["Urgent", "High", "Medium", "Low", "Routine"],
      REQUEST_TYPES: [
        "Data Pull", "Report Generation", "System Configuration",
        "Minor Bug Fix", "Content Update", "Training Material",
        "Documentation", "Analysis", "Process Improvement", "Other"
      ],
      DEFAULT_STATUS: "New",
      DEFAULT_PRIORITY: "Medium",
      // Drive folder ID for storing request attachments (must be pre-created)
      ATTACHMENTS_FOLDER_ID: "1Fyg1VyxIqFirz29uD0YLwpsKrTk5njRp",
    },

    // Organization configuration
    ORG: {
      // Node types for people (from Team List) and structural (derived at read time)
      PERSON_NODE_TYPES: ['director', 'deputy', 'lead', 'person', 'vacant'],
      STRUCTURAL_NODE_TYPES: ['hidden', 'task', 'team'],
      NODE_TYPES: ['hidden', 'director', 'deputy', 'task', 'team', 'lead', 'person', 'vacant'],
      // CPC format: 3 digits (LTR) — L=level, T=task index, R=role (reserved)
      // Digit 1 → node type mapping
      CPC_LEVEL_MAP: {
        '1': 'director',
        '2': 'deputy',
        '3': 'lead',
        '4': 'person'
      },
      DEFAULT_NODE_WIDTH: 200,
      DEFAULT_NODE_HEIGHT: 90,
      // Synthetic ID prefixes for structural nodes
      STRUCTURAL_ID_PREFIX: {
        TASK: 'task:',
        TEAM: 'team:',
        ROOT: 'root'
      }
    },

    // Notification settings
    NOTIFICATIONS: {
      ENABLED: true,
      DURATION: 3000, // milliseconds
      POLLING_INTERVAL: 60000, // 60 seconds for notification polling
      RETENTION_DAYS: null, // null = keep forever, set number for auto-cleanup
      DIGEST_ENABLED: false, // Phase 2: daily digest option
      TYPES: {
        SUCCESS: "success",
        ERROR: "error",
        INFO: "info",
        WARNING: "warning",
      },
    },

    // Pagination
    PAGINATION: {
      DEFAULT_PAGE_SIZE: 25,
      OPTIONS: [10, 25, 50, 100],
    },

    // Theme
    THEME: {
      DEFAULT: "light",
      OPTIONS: ["light", "dark"],
    },

    // Permissions
    PERMISSIONS: {
      ACTIONS: {
        VIEW: "view",
        EDIT: "edit",
        DELETE: "delete",
        ADMIN: "admin",
      },
    },
  };
})();
