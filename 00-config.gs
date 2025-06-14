/**
 * config.gs — Centralized Global Config & Constants
 * All IDs, sheet names, column names, secret codes, and config options used across the codebase.
 * Only the `config` object is global; everything else is scoped inside it.
 */

var config = {
  // USERS & CONTRACTORS SYSTEM
  USERS_SPREADSHEET_ID:      '1Mi0fOLNqzkPCp5yKQ7IJ_8nlVJRPUHd_QqNZkqFud90',
  USERS_SHEET_NAME:          'Users',
  CONTRACTORS_SPREADSHEET_ID:'1QknKTqJY98hQo50VKSHMmpQ4r9FAKvshun2t_dz32Ao',
  CONTRACTORS_SHEET_NAME:    'Contractors',
  DVI_CONTRACTORS_SHEET_NAME:  'Contractors',

  // Contractor column headers
  CONTRACTOR_ID_HEADER:      'ContractorID',
  CONTRACTOR_STATUS_HEADER:  'Status', // Value should be 'Active' for assignment

  // Internal user signup
  INTERNAL_SIGNUP_CODE:      'ADMIN2025', // CHANGE THIS as needed for prod security

  // Session & Security
  SESSION_TOKEN_LENGTH:      32,
  CACHE_EXPIRY_SECONDS:      60 * 60, // 1 hour

  // ROOT & LEAF PROJECT SOURCES
  PROJECT_SOURCES: {
    ROOT: {
      PROJECT_DATABASE_ID:    '193m8yOy51aDwSvqvCQQv7uM-fpiJrrnNK2kqCGYLq7E',
      TEMPLATE_SHEET_ID:      '1P5dcUNo62v4wg0n7I1dkLgBqV1Wh8g45qyWBqh3q6V8',
      DOC_TEMPLATE_ID:        '1q5ICoZswlYk1jC4dBC3HzmZP0OTdfQzYgV92dhm_UfI',
      PROJECTS_FOLDER_ID:     '1l2Ax8_urMaa7N90ZforGCSVkNsfs24_9'
    },
    LEAF: {
      PROJECT_DATABASE_ID:    '12yJhaOAe4rHSCFSadanh1K1YCu8wJnjbxfBdeX_GeB8',
      TEMPLATE_SHEET_ID:      '11Ysk5WLnSE5-Jri3Blh9SaiWmm6yAKi9ZseqdBsyjMk',
      DOC_TEMPLATE_ID:        '1GN2ELVOO0OKNeY0i0-UoM74XQgVdIIeQaQEKWc5YioY',
      PROJECTS_FOLDER_ID:     '1Sz3ju1QcrxWgAjpUe1sGkRJZZUgpAxS8'
    }
  },

  // Sheet names used in report generation
  SHEET_NAMES: {
    REPORT_INPUT:  'Report.Input',
    PROPOSED:      'Proposed',
    EXISTING:      'Existing',
    SYSTEMS:       'Systems',
    CHARTS:        'Charts'
  },

  // DVI / KANBAN / ASSIGNMENT
  DVI_PROJECTS_SPREADSHEET_ID: '1E64U9IQVlSoSNRv4X90BoH0JvonrZcZUXoV7a2GCjkU',
  DVI_PROJECTS_SHEET:          'dvi projects',

  // CHARTS & IMAGES
  CHARTS: {
    "<systems.graph.e>": "3.51848841E8",
    "<systems.graph.p>": "1.616441426E9"
  },
  PHOTOS: {
    PERCENT_WIDTH: 0.75,
    ASPECT_RATIO: 3 / 2
  },

  // FIELD MAPPINGS FOR PROJECT DB (indices)
  PROJECTID_COL:  10,
  REPORTID_COL:   10,
  STATUS_COL:     3,
  DOC_URL_COL:    1,
  SHEET_URL_COL:  0,

  // EMAIL CONFIG / NOTIFICATIONS
  SUPPORT_EMAIL: 'Support@RenewableEnergyIncentives.com',

  // WEBAPP DEPLOYMENT URL
  WEBAPP_DEPLOYMENT_URL: 'https://script.google.com/macros/s/AKfycbyOBD3TmfJrQyAsqlKvlOBk3-9cwOnWcOySCkgIaipFkXzibz9Eq1VV1Q7IhoIUbBU/exec'
  // Add any additional keys you want to expose here
};
Logger.log('SPREADSHEET ID:', config.CONTRACTORS_SPREADSHEET_ID);
Logger.log('SHEET NAME:', config.CONTRACTORS_SHEET_NAME);

function getContractorSheet() {
  return SpreadsheetApp.openById(config.CONTRACTORS_SPREADSHEET_ID)
                       .getSheetByName(config.CONTRACTORS_SHEET_NAME);
}
