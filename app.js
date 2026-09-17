/**
 * Freightoscope Bulk Import Leads Prototype Engine
 * Supports full Excel upload, parsing via SheetJS, 2-sheet template generation,
 * field and duplicate validation, permission toggling, and realistic UI state management.
 */

// Master Data Definitions
const COUNTRY_MASTER = {
  'US': 'United States',
  'IN': 'India',
  'AE': 'United Arab Emirates',
  'GB': 'United Kingdom',
  'SG': 'Singapore',
  'DE': 'Germany',
  'FR': 'France',
  'CN': 'China',
  'JP': 'Japan',
  'AU': 'Australia',
  'CA': 'Canada',
  'NL': 'Netherlands',
  'BR': 'Brazil',
  'IT': 'Italy',
  'ES': 'Spain',
  'MX': 'Mexico',
  'KR': 'South Korea',
  'SA': 'Saudi Arabia',
  'ZA': 'South Africa',
  'TR': 'Turkey',
  'DZ': 'Algeria',
  'MY': 'Malaysia',
  'VN': 'Vietnam',
  'TH': 'Thailand',
  'ID': 'Indonesia',
  'BD': 'Bangladesh',
  'DK': 'Denmark'
};

const USER_MASTER = [
  'sakshi.barnwal@freightoscope.com',
  'rafia.khan@freightoscope.com',
  'admin@freightoscope.com',
  'service@freightoscope.com',
  'john.doe@freightoscope.com',
  'anshul.p@freightoscope.com',
  'ahmed.raza@rgbx.com'
];

const LEAD_STAGES = [
  'New',
  'Contacted',
  'Qualified',
  'Proposal Sent',
  'Negotiation',
  'Won',
  'Lost'
];

// Active Parties for Party Lookup (Single-Select Lookup with Create)
const ACTIVE_PARTIES = [
  {
    name: 'Rajshahi Logistics',
    city: 'Rajshahi',
    country: 'Bangladesh',
    countryCode: 'BD',
    role: 'Lead',
    location: 'Rajshahi, Bangladesh'
  },
  {
    name: 'ABC Freight LLP Updated',
    city: 'New York',
    country: 'United States',
    countryCode: 'US',
    role: 'Customer',
    location: 'New York, United States'
  },
  {
    name: 'RGB Exports',
    city: 'Dubai',
    country: 'United Arab Emirates',
    countryCode: 'AE',
    role: 'Lead',
    location: 'Dubai, United Arab Emirates'
  },
  {
    name: 'Apex Global Logistics',
    city: 'Singapore',
    country: 'Singapore',
    countryCode: 'SG',
    role: 'Agent',
    location: 'Singapore, Singapore'
  },
  {
    name: 'Maersk Line',
    city: 'Copenhagen',
    country: 'Denmark',
    countryCode: 'DK',
    role: 'Customer',
    location: 'Copenhagen, Denmark'
  },
  {
    name: 'Global Freightways',
    city: 'London',
    country: 'United Kingdom',
    countryCode: 'GB',
    role: 'Lead',
    location: 'London, United Kingdom'
  }
];

// AC-5.2: Vendor Parties (Roles outside Lead/Agent/Customer excluded from lookup)
const VENDOR_PARTIES = [
  {
    name: 'Acme Corp',
    city: 'Chicago',
    country: 'United States',
    countryCode: 'US',
    role: 'Vendor'
  }
];

// AC-5.3: Inactive Parties (Excluded from lookup)
const INACTIVE_PARTIES = [
  {
    name: 'RGB Exports Inactive',
    city: 'Mumbai',
    country: 'India',
    countryCode: 'IN',
    role: 'Lead',
    active: false
  }
];

// Initial Contacts in Listing (Exact Match to Screenshot 1)
const INITIAL_CONTACTS = [
  {
    id: 'CNT-001',
    name: 'ABCDEFGHIJKLMNOPQRSTUVWXYABCDEFGHIJKLMNOPQRSTUVWXY',
    partyName: 'Rajshahi Logistics',
    partyLocation: 'Rajshahi, Bangladesh',
    owner: 'Rashed Saad',
    email: 'abcd@rajlog.com',
    phone: '--',
    createdBy: 'Rashed Saad',
    createdDate: '9/15/26, 3:00 PM'
  },
  {
    id: 'CNT-002',
    name: 'Dhruv Soni',
    partyName: 'ABC Freight LLP Updated',
    partyLocation: 'New York, United States',
    owner: 'Adish Chaudhari',
    email: 'dhruv1@fos.com',
    phone: '1122334455',
    createdBy: 'Adish Chaudhari',
    createdDate: '9/15/26, 11:34 AM'
  },
  {
    id: 'CNT-003',
    name: 'Sohel Chowdhury',
    partyName: 'Rajshahi Logistics',
    partyLocation: 'Rajshahi, Bangladesh',
    owner: 'Rashed Saad',
    email: 'sohel@rajlog.com',
    phone: '+88123456789',
    createdBy: 'Rashed Saad',
    createdDate: '9/8/26, 8:07 PM'
  },
  {
    id: 'CNT-004',
    name: 'Masud Rana',
    partyName: 'Rajshahi Logistics',
    partyLocation: 'Rajshahi, Bangladesh',
    owner: 'Rashed Saad',
    email: 'mr9@rajlog.com',
    phone: '0123456789',
    createdBy: 'Rashed Saad',
    createdDate: '9/8/26, 7:31 PM'
  }
];

let contactsData = JSON.parse(JSON.stringify(INITIAL_CONTACTS));
let activeModule = 'contacts'; // 'contacts' by default on load, or 'opportunities', 'leads'
let leadsCreatePermissionEnabled = true; // Controlled by Admin -> Roles -> Leads -> Add (AC-5.6)
let contactsCreatePermissionEnabled = true; // Controlled by Admin -> Roles -> Contact -> Add
let simulatePartySearchTimeout = false; // Toggleable for AC-5.5
let contactSearchQuery = '';

// Party Lookup-with-Create state:
let pendingNewParty = null; // { companyName, address, city, countryCode, countryName }
let selectedExistingParty = null;
let currentTypedPartyValue = '';
let lastAddressCapturePartyName = '';

// The 28 Predefined Column Headers
const DEFINED_HEADERS = [
  'Company Name*',
  'Short Name',
  'Legal Name',
  'Target Party Role',
  'Country Code*',
  'State',
  'City*',
  'Zip Code',
  'Address Line 1*',
  'Address Line 2',
  'Address Line 3',
  'Company Email',
  'Company Contact No.',
  'Company Website',
  'Source',
  'Lead Owner Email Id*',
  'Lead Stage*',
  'Industry',
  'Lead Temperature',
  'Service Interest',
  'Networks',
  'Notes',
  'Primary Contact',
  'First Name',
  'Last Name',
  'Email',
  'Phone No. (Primary)',
  'Designation'
];

// Initial Dummy Leads in Listing (Exact match to Screenshot 1)
let leadsData = [
  {
    leadId: 'RGBEXP',
    companyName: 'RGB Exports',
    location: 'New South Wales, Australia',
    stage: 'New',
    leadOwner: 'Ahmed Raza',
    temperature: '❄️',
    contactName: 'Mr. John Abel',
    contactPhone: '+61 412 345 678'
  },
  {
    leadId: 'GFA2026',
    companyName: 'GREAGR',
    location: 'Newyork, USA',
    stage: 'Contacted',
    leadOwner: 'Emily Carter',
    temperature: '🔥',
    contactName: 'Mrs. Britty Bell',
    contactPhone: '+1 412-555-0199'
  },
  {
    leadId: 'SBS2026',
    companyName: 'SkyBridge Solutions',
    location: 'Newyork, USA',
    stage: 'New',
    leadOwner: 'Ramesh Kumar',
    temperature: '❄️',
    contactName: 'Mr. Lukas Weber',
    contactPhone: '+1 412-555-0145'
  },
  {
    leadId: 'QES2026',
    companyName: 'QuantumEdge Systems',
    location: 'Newyork, USA',
    stage: 'Proposal Sent',
    leadOwner: 'Oliver Smith',
    temperature: '❄️',
    contactName: 'Mrs. Britty Bell',
    contactPhone: '+1 412-555-0182'
  },
  {
    leadId: 'BWL2026',
    companyName: 'BlueWave Logistics',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Lukas Weber',
    temperature: '☀️',
    contactName: 'Mr. Bob Abel',
    contactPhone: '+1 412-555-0123'
  },
  {
    leadId: 'NTX2026',
    companyName: 'Nexora Technologies',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Sarah Mitchell',
    temperature: '❄️',
    contactName: 'Mr. Bob Smith',
    contactPhone: '+1 412-555-0176'
  },
  {
    leadId: 'UNP2026',
    companyName: 'UrbanNest Properties',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Daniel Tan',
    temperature: '🔥',
    contactName: 'Mr. Bob Smith',
    contactPhone: '+1 412-555-0194'
  },
  {
    leadId: 'ESI2026',
    companyName: 'EcoSphere Innovations',
    location: 'Newyork, USA',
    stage: '--',
    leadOwner: 'Laura Bennett',
    temperature: '❄️',
    contactName: 'Mr. Khalid Hassan',
    contactPhone: '+1 412-555-0131'
  },
  {
    leadId: 'FINA2026',
    companyName: 'FinCore Analytics',
    location: 'Newyork, USA',
    stage: 'New',
    leadOwner: 'Khalid Hassan',
    temperature: '☀️',
    contactName: 'Mrs. Anjali Nair',
    contactPhone: '+1 412-555-0158'
  }
];

// State variables
let importPermissionEnabled = true;
let simulateBackendError = false;
let currentUploadedFile = null;
let parsedLeadRecords = [];
let validationErrors = [];
let isProcessing = false;

// DOM Elements
const drawerBackdrop = document.getElementById('drawerBackdrop');
const importDrawer = document.getElementById('importDrawer');
const btnImport = document.getElementById('btnImport');
const btnCloseDrawer = document.getElementById('btnCloseDrawer');
const btnDrawerCancel = document.getElementById('btnDrawerCancel');
const btnDrawerReset = document.getElementById('btnDrawerReset');
const btnDrawerSave = document.getElementById('btnDrawerSave');
const downloadTemplateLink = document.getElementById('downloadTemplateLink');

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const btnBrowseFiles = document.getElementById('btnBrowseFiles');
const uploadedFileBar = document.getElementById('uploadedFileBar');
const uploadedFileName = document.getElementById('uploadedFileName');
const uploadedFileSize = document.getElementById('uploadedFileSize');
const btnRemoveFile = document.getElementById('btnRemoveFile');

const bannerError = document.getElementById('bannerError');
const bannerErrorText = document.getElementById('bannerErrorText');
const errorsSection = document.getElementById('errorsSection');
const errorsTableBody = document.getElementById('errorsTableBody');
const errorCountPill = document.getElementById('errorCountPill');
const successValidationBox = document.getElementById('successValidationBox');
const successValidationText = document.getElementById('successValidationText');
const processingIndicator = document.getElementById('processingIndicator');

const leadsTableBody = document.getElementById('leadsTableBody');
const leadCountDisplay = document.getElementById('leadCountDisplay');
const importBtnWrapper = document.getElementById('importBtnWrapper');
const importTooltip = document.getElementById('importTooltip');

// Admin and Departments & Roles Submenu Elements
const menuAdmin = document.getElementById('menuAdmin');
const arrowAdmin = document.getElementById('arrowAdmin');
const submenuAdmin = document.getElementById('submenuAdmin');
const subitemDeptRoles = document.getElementById('subitemDeptRoles');

// Edit Roles & Permission Modal Elements
const rolesModalBackdrop = document.getElementById('rolesModalBackdrop');
const rolesModal = document.getElementById('rolesModal');
const btnCloseRolesModal = document.getElementById('btnCloseRolesModal');
const btnCancelRolesModal = document.getElementById('btnCancelRolesModal');
const btnSaveRolesModal = document.getElementById('btnSaveRolesModal');
const btnPermImport = document.getElementById('btnPermImport');
const wrapAllScreens = document.getElementById('wrapAllScreens');
const boxAllScreens = document.getElementById('boxAllScreens');
const searchScreensInput = document.getElementById('searchScreensInput');

// Sync Import Permission state across UI and button
function updateImportPermissionUI() {
  const btnPerm = document.getElementById('btnPermImport');
  if (btnPerm) {
    if (importPermissionEnabled) {
      btnPerm.classList.add('checked');
    } else {
      btnPerm.classList.remove('checked');
    }
  }

  if (importPermissionEnabled) {
    if (importBtnWrapper) importBtnWrapper.classList.remove('disabled');
    if (btnImport) {
      btnImport.classList.remove('disabled');
      btnImport.removeAttribute('disabled');
    }
  } else {
    if (importBtnWrapper) importBtnWrapper.classList.add('disabled');
    if (btnImport) {
      btnImport.classList.add('disabled');
      btnImport.setAttribute('disabled', 'true');
    }
  }
}

function toggleImportPermission(showToastNotice = false) {
  importPermissionEnabled = !importPermissionEnabled;
  updateImportPermissionUI();

  if (showToastNotice) {
    showToast(
      importPermissionEnabled ? 'Permission Granted' : 'Permission Revoked',
      importPermissionEnabled
        ? 'Import permission is ENABLED for Leads.'
        : 'Import permission is DISABLED for Leads.',
      importPermissionEnabled ? 'success' : 'warning',
      2500
    );
  }
}

function openRolesModal() {
  if (rolesModal && rolesModalBackdrop) {
    rolesModal.classList.add('open');
    rolesModalBackdrop.classList.add('open');
    updateImportPermissionUI();
  }
}

function closeRolesModal() {
  if (rolesModal && rolesModalBackdrop) {
    rolesModal.classList.remove('open');
    rolesModalBackdrop.classList.remove('open');
  }
}

// Helper to normalize header string: removes asterisks and dots, collapses spaces, trims
function cleanHeader(h) {
  if (!h) return '';
  return String(h)
    .replace(/[*_]/g, '')     // remove asterisks/underscores first
    .replace(/\s+/g, ' ')     // normalize multiple spaces to single space
    .trim()                   // trim leading and trailing spaces AFTER removing symbols!
    .toLowerCase();           // lowercase for case-insensitive matching
}

// Map clean headers to standard keys
const HEADER_KEY_MAP = {
  'company name': 'companyName',
  'short name': 'shortName',
  'legal name': 'legalName',
  'target party role': 'targetPartyRole',
  'country code': 'countryCode',
  'state': 'state',
  'city': 'city',
  'zip code': 'zipCode',
  'address line 1': 'addressLine1',
  'address line 2': 'addressLine2',
  'address line 3': 'addressLine3',
  'company email': 'companyEmail',
  'company contact no.': 'companyContactNo',
  'company contact no': 'companyContactNo',
  'company website': 'companyWebsite',
  'source': 'source',
  'lead owner email id': 'leadOwnerEmailId',
  'lead stage': 'leadStage',
  'industry': 'industry',
  'lead temperature': 'leadTemperature',
  'service interest': 'serviceInterest',
  'networks': 'networks',
  'notes': 'notes',
  'primary contact': 'primaryContact',
  'first name': 'firstName',
  'last name': 'lastName',
  'email': 'email',
  'phone no. (primary)': 'phonePrimary',
  'phone no (primary)': 'phonePrimary',
  'phone no.': 'phonePrimary',
  'designation': 'designation'
};

// Initial Render
document.addEventListener('DOMContentLoaded', () => {
  renderLeadsTable();
  updateImportPermissionUI();
  initOpportunitiesModule();
  initContactsModule();
  setupEventListeners();
  showContactsModule();
});

function setupEventListeners() {
  // Admin Submenu Toggle
  if (menuAdmin && submenuAdmin) {
    menuAdmin.addEventListener('click', () => {
      const isCurrentlyOpen = submenuAdmin.style.display === 'block';
      submenuAdmin.style.display = isCurrentlyOpen ? 'none' : 'block';
      if (arrowAdmin) {
        arrowAdmin.style.transform = isCurrentlyOpen ? 'rotate(0deg)' : 'rotate(90deg)';
      }
    });
  }

  // Departments & Roles Submenu Item Click -> Opens Edit Roles & Permission Modal
  if (subitemDeptRoles) {
    subitemDeptRoles.addEventListener('click', (e) => {
      e.stopPropagation();
      openRolesModal();
    });
  }

  // Modal Close, Cancel, and Backdrop Handlers
  if (btnCloseRolesModal) btnCloseRolesModal.addEventListener('click', closeRolesModal);
  if (btnCancelRolesModal) btnCancelRolesModal.addEventListener('click', closeRolesModal);
  if (rolesModalBackdrop) rolesModalBackdrop.addEventListener('click', closeRolesModal);

  // Modal Import Permission Button Toggle
  if (btnPermImport) {
    btnPermImport.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleImportPermission(true);
    });
  }

  // Modal Sales Opportunity Update Permission Button Toggle
  const btnPermOppUpdate = document.getElementById('btnPermOppUpdate');
  if (btnPermOppUpdate) {
    btnPermOppUpdate.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleOppUpdatePermission();
    });
  }

  // Modal Leads Add/Create Permission Toggle (AC-5.6)
  const btnPermLeadAdd = document.getElementById('btnPermLeadAdd');
  if (btnPermLeadAdd) {
    btnPermLeadAdd.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleLeadsCreatePermission(true);
    });
  }

  // Modal Contact Add Permission Toggle
  const btnPermContactAdd = document.getElementById('btnPermContactAdd');
  if (btnPermContactAdd) {
    btnPermContactAdd.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleContactsCreatePermission(true);
    });
  }

  // Modal Save Button
  if (btnSaveRolesModal) {
    btnSaveRolesModal.addEventListener('click', () => {
      updateImportPermissionUI();
      saveOppPermissions();
      closeRolesModal();
      showToast(
        'Permissions Saved',
        `Permissions updated successfully. Leads Create permission is ${leadsCreatePermissionEnabled ? 'ENABLED' : 'DISABLED'}.`,
        leadsCreatePermissionEnabled ? 'success' : 'warning',
        4500
      );
    });
  }

  // Module items click selection
  document.querySelectorAll('.module-item').forEach(item => {
    item.addEventListener('click', () => {
      document.querySelectorAll('.module-item').forEach(m => m.classList.remove('active'));
      item.classList.add('active');
    });
  });

  // Action pills toggle (Add, Delete, Update, View)
  document.querySelectorAll('.btn-action-pill').forEach(pill => {
    if (pill.id === 'btnPermImport' || pill.id === 'btnPermOppUpdate' || pill.id === 'btnPermLeadAdd' || pill.id === 'btnPermContactAdd') return; // Handled exclusively
    pill.addEventListener('click', (e) => {
      e.preventDefault();
      pill.classList.toggle('checked');
    });
  });

  // Screen row left checkbox toggle
  document.querySelectorAll('.screen-row-left').forEach(rowLeft => {
    rowLeft.addEventListener('click', (e) => {
      e.preventDefault();
      const box = rowLeft.querySelector('.screen-check-box');
      if (box) {
        box.classList.toggle('checked');
        const isChecked = box.classList.contains('checked');
        const row = rowLeft.closest('.screen-row');
        if (row) {
          row.querySelectorAll('.btn-action-pill').forEach(pill => {
            pill.classList.toggle('checked', isChecked);
          });
          if (row.id === 'rowLeadsScreen') {
            importPermissionEnabled = isChecked;
            updateImportPermissionUI();
          }
        }
      }
    });
  });

  // All Screens Toggle
  if (wrapAllScreens && boxAllScreens) {
    wrapAllScreens.addEventListener('click', (e) => {
      e.preventDefault();
      boxAllScreens.classList.toggle('checked');
      const allChecked = boxAllScreens.classList.contains('checked');
      document.querySelectorAll('.screen-check-box').forEach(b => {
        b.classList.toggle('checked', allChecked);
      });
      document.querySelectorAll('.btn-action-pill').forEach(p => {
        p.classList.toggle('checked', allChecked);
      });
      importPermissionEnabled = allChecked;
      updateImportPermissionUI();
    });
  }

  // Search screens input
  if (searchScreensInput) {
    searchScreensInput.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      document.querySelectorAll('.screens-list-container .screen-row').forEach(row => {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(q) ? 'flex' : 'none';
      });
    });
  }


  // Import Button & Wrapper Click (Enforces Permission & Tooltip)
  if (importBtnWrapper) {
    importBtnWrapper.addEventListener('click', (e) => {
      if (!importPermissionEnabled) {
        // Flash tooltip
        importBtnWrapper.classList.add('show-tooltip');
        setTimeout(() => importBtnWrapper.classList.remove('show-tooltip'), 3000);

        showToast(
          'Access Restricted',
          'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance',
          'error',
          6500
        );
        return;
      }
      openDrawer();
    });
  }

  if (btnImport) {
    btnImport.addEventListener('click', (e) => {
      if (!importPermissionEnabled) {
        e.preventDefault();
        e.stopPropagation();
        importBtnWrapper.classList.add('show-tooltip');
        setTimeout(() => importBtnWrapper.classList.remove('show-tooltip'), 3000);
        showToast(
          'Access Restricted',
          'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance',
          'error',
          6500
        );
        return;
      }
      openDrawer();
    });
  }

  // Close Drawer
  btnCloseDrawer.addEventListener('click', closeDrawer);
  btnDrawerCancel.addEventListener('click', closeDrawer);
  drawerBackdrop.addEventListener('click', closeDrawer);

  // Download Template (generates 2 sheets)
  downloadTemplateLink.addEventListener('click', (e) => {
    e.preventDefault();
    generateAndDownloadTemplate();
  });

  // Browse files button
  btnBrowseFiles.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleSelectedFile(e.target.files[0]);
    }
  });

  // Drag and Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('drag-over');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('drag-over');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) {
      handleSelectedFile(dt.files[0]);
    }
  });

  // Remove File
  btnRemoveFile.addEventListener('click', (e) => {
    e.stopPropagation();
    resetUploadSection();
  });

  // Reset Button
  btnDrawerReset.addEventListener('click', () => {
    resetUploadSection();
  });

  // Save Button
  btnDrawerSave.addEventListener('click', handleFinalSave);
}

// Drawer visibility controls
function openDrawer() {
  importDrawer.classList.add('open');
  drawerBackdrop.classList.add('open');
}

function closeDrawer() {
  if (isProcessing) {
    // Abort processing as specified: "If a file is currently Processing, clicking close must abort the operation."
    isProcessing = false;
    hideProcessing();
  }
  importDrawer.classList.remove('open');
  drawerBackdrop.classList.remove('open');
}

// Reset upload section
function resetUploadSection() {
  currentUploadedFile = null;
  fileInput.value = '';
  parsedLeadRecords = [];
  validationErrors = [];
  isProcessing = false;

  uploadedFileBar.style.display = 'none';
  dropzone.style.display = 'flex';
  btnDrawerReset.style.display = 'none';
  hideBanners();
  hideErrors();
  hideSuccessBox();
  hideProcessing();

  setSaveEnabled(false);
}

function setSaveEnabled(enabled) {
  if (enabled) {
    btnDrawerSave.disabled = false;
    btnDrawerSave.classList.add('active');
  } else {
    btnDrawerSave.disabled = true;
    btnDrawerSave.classList.remove('active');
  }
}

function showProcessing() {
  isProcessing = true;
  processingIndicator.style.display = 'flex';
}

function hideProcessing() {
  isProcessing = false;
  processingIndicator.style.display = 'none';
}

function showBannerError(msg) {
  bannerErrorText.textContent = msg;
  bannerError.style.display = 'block';
}

function hideBanners() {
  bannerError.style.display = 'none';
  bannerErrorText.textContent = '';
}

function hideErrors() {
  errorsSection.style.display = 'none';
  errorsTableBody.innerHTML = '';
}

function hideSuccessBox() {
  successValidationBox.style.display = 'none';
  successValidationText.textContent = '';
}

// File Validation & Handling
function handleSelectedFile(file) {
  hideBanners();
  hideErrors();
  hideSuccessBox();
  setSaveEnabled(false);

  const fileName = file.name;
  const ext = fileName.slice((fileName.lastIndexOf(".") - 1 >>> 0) + 2).toLowerCase();

  // 1. File Format Restriction: Accept only .xls and .xlsx
  if (ext !== 'xlsx' && ext !== 'xls') {
    // Other formats are rejected, not displayed, and show: "Invalid format found. Please Select a valid format"
    showBannerError('Invalid format found. Please Select a valid format');
    fileInput.value = '';
    return;
  }

  // Display file in blue bar below upload area
  currentUploadedFile = file;
  uploadedFileName.textContent = file.name;
  uploadedFileSize.textContent = formatBytes(file.size);
  uploadedFileBar.style.display = 'flex';
  btnDrawerReset.style.display = 'inline-block';

  showProcessing();

  // Parse Excel via SheetJS
  const reader = new FileReader();
  reader.onload = (e) => {
    if (!isProcessing) return; // aborted
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      processWorkbook(workbook);
    } catch (err) {
      hideProcessing();
      showBannerError('Failed to parse the file. Please ensure it is a valid Excel spreadsheet.');
    }
  };
  reader.onerror = () => {
    hideProcessing();
    showBannerError('Error reading file. Please try again.');
  };
  reader.readAsArrayBuffer(file);
}

function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Process Workbook and Validate
function processWorkbook(workbook) {
  hideProcessing();

  // Find sheet: look for sheet containing lead data, ignoring instructions
  let sheetName = workbook.SheetNames.find(n =>
    (n.toLowerCase().includes('bulk') || n.toLowerCase().includes('lead') || n.toLowerCase().includes('import')) &&
    !n.toLowerCase().includes('instruction')
  );

  if (!sheetName) {
    for (const name of workbook.SheetNames) {
      if (name.toLowerCase().includes('instruction')) continue;
      sheetName = name;
      break;
    }
  }

  if (!sheetName) {
    sheetName = workbook.SheetNames[0];
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    showBannerError('No valid sheet found in uploaded Excel file.');
    return;
  }

  // Read as array of arrays
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (!rawData || rawData.length === 0) {
    showBannerError('No lead found in the uploaded file. Please add at least one lead detail and upload again.');
    return;
  }

  // Header row (Row 1)
  const headerRow = rawData[0];
  if (!headerRow || headerRow.length === 0) {
    showBannerError('Incorrect column names found. Please upload a correct excel file');
    return;
  }

  // 2. File Parsing & Header Validation
  // Case-insensitively matched, column order irrelevant.
  // Missing columns shall block import and display: "Incorrect column names found. Please upload a correct excel file"
  // For duplicate columns, only first occurrence considered, subsequent ignored and not overwrite.

  const columnMap = {}; // cleanHeader -> first col index
  const foundCleanHeaders = new Set();

  headerRow.forEach((colName, colIdx) => {
    const clean = cleanHeader(colName);
    if (clean && !foundCleanHeaders.has(clean)) {
      foundCleanHeaders.add(clean);
      columnMap[clean] = colIdx;
    }
  });

  // Verify all 28 defined headers are present
  let missingRequiredTemplateHeaders = false;
  for (const defHeader of DEFINED_HEADERS) {
    const clean = cleanHeader(defHeader);
    if (!foundCleanHeaders.has(clean)) {
      missingRequiredTemplateHeaders = true;
      break;
    }
  }

  if (missingRequiredTemplateHeaders) {
    showBannerError('Incorrect column names found. Please upload a correct excel file');
    return;
  }

  // 3. Extract Rows & Empty file check
  const rows = rawData.slice(1);
  const records = [];

  rows.forEach((rowArr, index) => {
    const rowNum = index + 2; // Excel row number (1-based, Row 1 is header)
    
    // Check if entire row is empty
    const isRowEmpty = rowArr.every(val => val === '' || val === null || val === undefined || String(val).trim() === '');
    if (!isRowEmpty) {
      const record = { _excelRow: rowNum };
      for (const defHeader of DEFINED_HEADERS) {
        const clean = cleanHeader(defHeader);
        const colIdx = columnMap[clean];
        const val = (colIdx !== undefined && rowArr[colIdx] !== undefined) ? String(rowArr[colIdx]).trim() : '';
        const key = HEADER_KEY_MAP[clean] || clean;
        record[key] = val;
      }
      records.push(record);
    }
  });

  // "If the uploaded Excel file contains no Lead details and all Lead fields are blank, the system shall display:
  // 'No lead found in the uploaded file. Please add at least one lead detail and upload again.'"
  if (records.length === 0) {
    showBannerError('No lead found in the uploaded file. Please add at least one lead detail and upload again.');
    return;
  }

  parsedLeadRecords = records;

  // 4. Run Field Validations and Duplicate Validations
  runValidationSuite(records);
}

// Check if an email belongs to FOS (Freightoscope internal domain)
function isFosEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const trimmed = email.trim().toLowerCase();
  const atIdx = trimmed.lastIndexOf('@');
  if (atIdx === -1 || atIdx === 0 || atIdx === trimmed.length - 1) return false;
  const domain = trimmed.slice(atIdx + 1);
  if (!domain) return false;

  // Exact FOS domains or subdomains (*.fos.com, *.freightoscope.com)
  if (domain === 'fos.com' || domain.endsWith('.fos.com')) return true;
  if (domain === 'freightoscope.com' || domain.endsWith('.freightoscope.com')) return true;

  // Also support any regional FOS domain (e.g. fos.in, freightoscope.ae)
  if (/^([a-zA-Z0-9-]+\.)*(fos|freightoscope)\.[a-zA-Z]{2,}$/i.test(domain)) return true;

  return false;
}

// Validate company website URL format
function isValidCompanyWebsite(urlStr) {
  if (!urlStr || typeof urlStr !== 'string') return false;
  const str = urlStr.trim();
  if (!str) return false;

  // Reject if contains whitespace
  if (/\s/.test(str)) return false;

  // Reject typos of protocol like "https." or "http." or "http:/" or "https:/" or "http//"
  if (/^https?[\.:\/]/i.test(str) && !/^https?:\/\//i.test(str)) {
    return false;
  }

  // Reject other protocols (ftp:, mailto:, file:, javascript:)
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(str) && !/^https?:\/\//i.test(str)) {
    return false;
  }

  // Strip http:// or https:// for domain parsing
  let withoutProtocol = str.replace(/^https?:\/\//i, '');

  // Separate path/query/hash from host
  const slashIdx = withoutProtocol.indexOf('/');
  const questionIdx = withoutProtocol.indexOf('?');
  const hashIdx = withoutProtocol.indexOf('#');
  let cutIdx = withoutProtocol.length;
  if (slashIdx !== -1 && slashIdx < cutIdx) cutIdx = slashIdx;
  if (questionIdx !== -1 && questionIdx < cutIdx) cutIdx = questionIdx;
  if (hashIdx !== -1 && hashIdx < cutIdx) cutIdx = hashIdx;

  const hostPort = withoutProtocol.slice(0, cutIdx);
  const host = hostPort.split(':')[0]; // strip optional port

  if (!host || !host.includes('.')) return false;
  if (host.startsWith('.') || host.endsWith('.') || host.startsWith('-') || host.endsWith('-')) return false;

  const parts = host.split('.');
  if (parts.length < 2) return false;

  // Subdomain cannot be "http" or "https" (e.g. "https.dbjbf.com")
  if (parts[0].toLowerCase() === 'http' || parts[0].toLowerCase() === 'https') {
    return false;
  }

  // Verify each label
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part || part.length === 0) return false;
    if (!/^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?$/.test(part)) {
      return false;
    }
  }

  // TLD must be at least 2 alphabetical characters
  const tld = parts[parts.length - 1];
  if (!/^[a-zA-Z]{2,}$/.test(tld)) {
    return false;
  }

  return true;
}

// Validation Suite
function runValidationSuite(records) {
  validationErrors = [];

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const iso2Regex = /^[a-zA-Z]{2}$/;
  const phoneLetterRegex = /[a-zA-Z]/;

  // Track duplicates: Company Name + City + Country (normalized)
  const duplicateLeadMap = new Map();

  records.forEach((rec, idx) => {
    const row = rec._excelRow;

    // Company Name
    // Empty -> "Company Name is missing"
    // > 100 chars -> "Company Name cannot exceed 100 characters"
    if (!rec.companyName) {
      addError(row, 'Company Name', 'Missing Field', 'Company Name is missing');
    } else if (rec.companyName.length > 100) {
      addError(row, 'Company Name', 'Character Limit Exceeded', 'Company Name cannot exceed 100 characters');
    }

    // Country Code
    // Empty -> "Country Code is missing"
    // Not valid 2-letter ISO code -> "Country Code must be a valid 2-letter ISO code"
    // Does not match existing Country -> "This country could not be found"
    if (!rec.countryCode) {
      addError(row, 'Country Code', 'Missing Field', 'Country Code is missing');
    } else if (!iso2Regex.test(rec.countryCode)) {
      addError(row, 'Country Code', 'Invalid Format', 'Country Code must be a valid 2-letter ISO code');
    } else if (!COUNTRY_MASTER[rec.countryCode.toUpperCase()]) {
      addError(row, 'Country Code', 'Country Not Found', 'This country could not be found');
    }

    // City
    // Empty -> "City is missing"
    // > 100 chars -> "City cannot exceed 100 characters"
    if (!rec.city) {
      addError(row, 'City', 'Missing Field', 'City is missing');
    } else if (rec.city.length > 100) {
      addError(row, 'City', 'Character Limit Exceeded', 'City cannot exceed 100 characters');
    }

    // Address Line 1
    // Empty -> "Address Line 1 is missing"
    // > 500 chars -> "Address Line 1 cannot exceed 500 characters"
    if (!rec.addressLine1) {
      addError(row, 'Address Line 1', 'Missing Field', 'Address Line 1 is missing');
    } else if (rec.addressLine1.length > 500) {
      addError(row, 'Address Line 1', 'Character Limit Exceeded', 'Address Line 1 cannot exceed 500 characters');
    }

    // Company Email
    if (rec.companyEmail) {
      const cEmail = rec.companyEmail.trim();
      if (!emailRegex.test(cEmail)) {
        addError(row, 'Company Email', 'Invalid Format', 'This company email address is not valid');
      } else if (cEmail.length > 100) {
        addError(row, 'Company Email', 'Character Limit Exceeded', 'Company email cannot exceed 100 characters');
      }
    }

    // Company Contact No.
    if (rec.companyContactNo) {
      const cPhone = rec.companyContactNo.trim();
      if (cPhone.length > 20) {
        addError(row, 'Company Contact No.', 'Character Limit Exceeded', 'Company contact number cannot exceed 20 characters');
      } else if (phoneLetterRegex.test(cPhone)) {
        addError(row, 'Company Contact No.', 'Invalid Format', 'Company contact number can only contain digits & special characters');
      }
    }

    // Company Website
    // If entered:
    // > 100 chars -> "Company website cannot exceed 100 characters"
    // Invalid format -> "This company website is not valid"
    if (rec.companyWebsite) {
      const site = rec.companyWebsite.trim();
      if (site.length > 100) {
        addError(row, 'Company Website', 'Character Limit Exceeded', 'Company website cannot exceed 100 characters');
      } else if (!isValidCompanyWebsite(site)) {
        addError(row, 'Company Website', 'Invalid Format', 'This company website is not valid');
      }
    }

    // Lead Owner Email Id
    // Empty -> "Lead Owner Email Id is missing"
    // > 100 chars -> "Lead Owner Email Id cannot exceed 100 characters"
    // Invalid email format -> "This owner email address is not valid"
    // Not an FOS mail -> "This lead owner could not be found"
    const ownerEmail = (rec.leadOwnerEmailId || '').trim();
    if (!ownerEmail) {
      addError(row, 'Lead Owner Email Id', 'Missing Field', 'Lead Owner Email Id is missing');
    } else if (ownerEmail.length > 100) {
      addError(row, 'Lead Owner Email Id', 'Character Limit Exceeded', 'Lead Owner Email Id cannot exceed 100 characters');
    } else if (!emailRegex.test(ownerEmail)) {
      addError(row, 'Lead Owner Email Id', 'Invalid Format', 'This owner email address is not valid');
    } else if (!isFosEmail(ownerEmail)) {
      addError(row, 'Lead Owner Email Id', 'Owner Not Found', 'This lead owner could not be found');
    }

    // Lead Stage
    // Empty -> "Lead Stage is missing"
    if (!rec.leadStage) {
      addError(row, 'Lead Stage', 'Missing Field', 'Lead Stage is missing');
    }

    // Primary Contact: must be TRUE, FALSE, or blank
    let primaryContactRaw = rec.primaryContact !== undefined && rec.primaryContact !== null ? String(rec.primaryContact).trim() : '';
    let isPrimaryTrue = false;
    let isPrimaryFalse = false;

    if (primaryContactRaw === '1' || primaryContactRaw.toUpperCase() === 'TRUE') {
      isPrimaryTrue = true;
    } else if (primaryContactRaw === '0' || primaryContactRaw.toUpperCase() === 'FALSE' || primaryContactRaw === '') {
      isPrimaryFalse = true;
    } else {
      addError(row, 'Primary Contact', 'Invalid Value', 'Primary Contact must be TRUE or FALSE');
    }

    // Primary Contact = TRUE validations
    if (isPrimaryTrue) {
      // First Name
      if (!rec.firstName) {
        addError(row, 'First Name', 'Missing Field', 'First Name is missing (required when Primary Contact is TRUE)');
      } else if (rec.firstName.length > 50) {
        addError(row, 'First Name', 'Character Limit Exceeded', 'First Name cannot exceed 50 characters');
      }

      // Last Name
      if (!rec.lastName) {
        addError(row, 'Last Name', 'Missing Field', 'Last Name is missing (required when Primary Contact is TRUE)');
      } else if (rec.lastName.length > 50) {
        addError(row, 'Last Name', 'Character Limit Exceeded', 'Last Name cannot exceed 50 characters');
      }

      // Email
      if (!rec.email) {
        addError(row, 'Email', 'Missing Field', 'Email is missing (required when Primary Contact is TRUE)');
      } else if (!emailRegex.test(rec.email)) {
        addError(row, 'Email', 'Invalid Format', 'This email address is not valid');
      } else if (rec.email.length > 100) {
        addError(row, 'Email', 'Character Limit Exceeded', 'Email cannot exceed 100 characters');
      }

      // Phone No. (Primary)
      if (rec.phonePrimary) {
        if (phoneLetterRegex.test(rec.phonePrimary)) {
          addError(row, 'Phone No. (Primary)', 'Invalid Format', 'Phone number can only contain digits & special characters');
        } else {
          // Count digits
          const digitCount = (rec.phonePrimary.match(/\d/g) || []).length;
          if (digitCount > 20) {
            addError(row, 'Phone No. (Primary)', 'Character Limit Exceeded', 'Phone number cannot exceed 20 digits');
          }
        }
      }
    }

    // Build Duplicate Key (Company Name + City + Country)
    const normComp = (rec.companyName || '').trim().toLowerCase();
    const normCity = (rec.city || '').trim().toLowerCase();
    const normCountry = (rec.countryCode || '').trim().toLowerCase();
    
    if (normComp && normCity && normCountry) {
      const dupKey = `${normComp}###${normCity}###${normCountry}`;
      if (!duplicateLeadMap.has(dupKey)) {
        duplicateLeadMap.set(dupKey, [row]);
      } else {
        duplicateLeadMap.get(dupKey).push(row);
      }
    }
  });

  // Check Duplicate Entries
  for (const [key, rowList] of duplicateLeadMap.entries()) {
    if (rowList.length > 1) {
      // Multiple rows share the same Company Name, City, Country
      rowList.forEach(rNum => {
        addError(
          rNum,
          'Company Name + City + Country',
          'Duplicate Entry',
          'This lead is already present in the uploaded file'
        );
      });
    }
  }

  // Display Results
  displayValidationSummary(records.length);
}

function addError(row, col, type, message) {
  validationErrors.push({ row, col, type, message });
}

function displayValidationSummary(totalRecords) {
  if (validationErrors.length > 0) {
    // Sort errors by Excel Row then Column
    validationErrors.sort((a, b) => a.row - b.row);

    errorCountPill.textContent = `${validationErrors.length} ${validationErrors.length === 1 ? 'Error' : 'Errors'}`;
    errorsTableBody.innerHTML = '';

    validationErrors.forEach(err => {
      const tr = document.createElement('tr');

      let badgeClass = 'format';
      if (err.type === 'Missing Field') badgeClass = 'missing';
      else if (err.type === 'Character Limit Exceeded') badgeClass = 'limit';
      else if (err.type === 'Duplicate Entry') badgeClass = 'duplicate';
      else if (err.type === 'Country Not Found' || err.type === 'Owner Not Found') badgeClass = 'notfound';
      else if (err.type === 'Invalid Value') badgeClass = 'invalid';

      tr.innerHTML = `
        <td style="font-weight: 700; color: #1e293b;">${err.row}</td>
        <td style="font-weight: 600; color: #334155;">${escapeHtml(err.col)}</td>
        <td><span class="error-badge ${badgeClass}">${escapeHtml(err.type)}</span></td>
        <td style="color: #b42318; font-weight: 500;">${escapeHtml(err.message)}</td>
      `;
      errorsTableBody.appendChild(tr);
    });

    errorsSection.style.display = 'flex';
    hideSuccessBox();
    setSaveEnabled(false);
  } else {
    // All rows passed validation
    hideErrors();
    successValidationText.textContent = `All rows passed validation. ${totalRecords} lead(s) are ready to import.`;
    successValidationBox.style.display = 'flex';
    setSaveEnabled(true);
  }
}

// Final Save Processing
function handleFinalSave() {
  if (validationErrors.length > 0 || parsedLeadRecords.length === 0) return;

  // Check Backend Error Simulation
  if (simulateBackendError) {
    showToast(
      'Import Failed',
      'Failed to import leads. Please try again',
      'error',
      6000
    );
    // As per requirement: "The drawer remains open"
    return;
  }

  // Generate unique Lead IDs and Map Primary Contacts
  const newLeads = parsedLeadRecords.map((rec, idx) => {
    // Lead ID Generation: prefix from company name + random/seq
    const cleanComp = rec.companyName.replace(/[^a-zA-Z]/g, '').toUpperCase().slice(0, 6) || 'LEAD';
    const seq = Math.floor(100 + Math.random() * 900);
    const generatedLeadId = `${cleanComp}${seq}`;

    // Country Display
    const countryName = COUNTRY_MASTER[(rec.countryCode || '').toUpperCase()] || rec.countryCode;
    const locationStr = rec.state ? `${rec.city}, ${rec.state}, ${countryName}` : `${rec.city}, ${countryName}`;

    // Primary Contact mapping
    let primaryContactDisplay = '--';
    const pcStr = (rec.primaryContact || '').toString().trim().toUpperCase();
    if (pcStr === '1' || pcStr === 'TRUE') {
      primaryContactDisplay = `${rec.firstName} ${rec.lastName}`.trim();
    }

    // Lead Owner display name
    let leadOwnerDisplay = rec.leadOwnerEmailId;
    if (rec.leadOwnerEmailId.includes('sakshi')) leadOwnerDisplay = 'Sakshi Barnwal';
    else if (rec.leadOwnerEmailId.includes('rafia')) leadOwnerDisplay = 'Rafia Khan';
    else if (rec.leadOwnerEmailId.includes('john')) leadOwnerDisplay = 'John Doe';
    else if (rec.leadOwnerEmailId.includes('anshul')) leadOwnerDisplay = 'Anshul P';
    else if (rec.leadOwnerEmailId.includes('admin')) leadOwnerDisplay = 'Admin';

    return {
      leadId: generatedLeadId,
      companyName: rec.companyName,
      location: locationStr,
      stage: rec.leadStage || 'New',
      temperature: rec.leadTemperature === 'Warm' ? '☀️' : (rec.leadTemperature === 'Cold' ? '❄️' : '🔥'),
      leadOwner: leadOwnerDisplay,
      contactName: primaryContactDisplay,
      contactPhone: rec.phonePrimary || '--',
      isNew: true
    };
  });

  // Prepend new leads to table
  leadsData = [...newLeads, ...leadsData];
  renderLeadsTable();

  // Close Drawer
  closeDrawer();
  resetUploadSection();

  // Success toast
  showToast(
    'Success',
    'Leads imported successfully',
    'success'
  );
}

// Render Leads Listing Table (Exact match to Screenshot 1)
function renderLeadsTable() {
  if (leadsTableBody) leadsTableBody.innerHTML = '';
  if (leadCountDisplay) leadCountDisplay.textContent = leadsData.length;

  leadsData.forEach(lead => {
    const tr = document.createElement('tr');
    if (lead.isNew) tr.classList.add('newly-imported');

    let stageBadge = `<span class="badge-new-solid">${escapeHtml(lead.stage)}</span>`;
    if (lead.stage === 'Contacted') {
      stageBadge = `<span class="badge-contacted-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === 'Proposal Sent') {
      stageBadge = `<span class="badge-proposal-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === 'Qualified') {
      stageBadge = `<span class="badge-qualified-solid">${escapeHtml(lead.stage)}</span>`;
    } else if (lead.stage === '--') {
      stageBadge = `<span style="color: #9ca3af; font-weight: 500;">--</span>`;
    }

    tr.innerHTML = `
      <td style="width: 32px; text-align: center;"></td>
      <td class="icon-cell">⋮</td>
      <td class="icon-cell" style="color: #d1d5db;">☆</td>
      <td><a class="lead-id-link">${escapeHtml(lead.leadId)}</a></td>
      <td>
        <div class="lead-company-name">${escapeHtml(lead.companyName)}</div>
        <div class="lead-location-sub">${escapeHtml(lead.location)}</div>
      </td>
      <td>${stageBadge}</td>
      <td>${escapeHtml(lead.leadOwner)}</td>
      <td style="font-size: 14px; text-align: center;">${lead.temperature}</td>
      <td>${escapeHtml(lead.contactName || '--')}</td>
      <td>${escapeHtml(lead.contactPhone || '--')}</td>
    `;
    leadsTableBody.appendChild(tr);
  });
}

// Generate & Download Real 2-Sheet Excel Template
function generateAndDownloadTemplate() {
  if (typeof XLSX === 'undefined') {
    showToast('Error', 'SheetJS library not loaded. Please ensure xlsx.full.min.js is present.', 'error');
    return;
  }

  const wb = XLSX.utils.book_new();

  // 1. Sheet 1: "Import Bulk Lead"
  const sheet1Data = [
    DEFINED_HEADERS,
    [
      'Acme Global Logistics',
      'AGL',
      'Acme Global Logistics Pvt. Ltd.',
      'Customer',
      'US',
      'Texas',
      'Dallas',
      '75001',
      '1200 Logistics Blvd, Suite 400',
      'Building B',
      '',
      'info@acme-globallog.com',
      '+1 214-555-0199',
      'www.acme-globallog.com',
      'Direct / Inbound',
      'sakshi.barnwal@freightoscope.com',
      'New',
      'Logistics',
      'Hot',
      'Ocean Freight (FCL)',
      'WCA',
      'Referred from partner network. Interested in FCL contracts.',
      'TRUE',
      'Sharon',
      'Roy',
      'sharon.roy@acme-globallog.com',
      '+1 214-555-0145',
      'Logistics Manager'
    ]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);

  // Auto-width for Sheet 1
  ws1['!cols'] = DEFINED_HEADERS.map(h => ({ wch: Math.max(h.length + 3, 16) }));

  // 2. Sheet 2: "Instructions"
  const instructionHeaders = [
    'Field',
    'Field Value',
    'No of Characters & Type',
    'Mandatory/Optional',
    'Other Validations'
  ];

  const instructionRows = [
    ['Company Name*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. Company Name must have a value.\n2. Value must not exceed 100 characters.'],
    ['Short Name', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 100 characters.'],
    ['Legal Name', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 100 characters.'],
    ['Target Party Role', 'Dropdown (Hardcoded)', 'Value from Dropdown (Agent, Customer)', 'Optional', '1. Each cell of the column contains the Target Party Role options: Agent, Customer.\n2. One of the options should be selected from the dropdown.\n3. If no option is selected or an incorrect value is entered, Target Party Role shall be empty for the Lead.'],
    ['Country Code*', 'Lookup', '2 Alpha Characters (ISO Code)', 'Mandatory', '1. Country Code must have a value.\n2. Value must be a valid 2-letter ISO Country Code.\n3. Value must match an existing Country. Source: Master -> Locations -> Countries.'],
    ['State', 'Lookup (Free text fallback)', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If the value matches an existing State (under the selected Country) in Master -> Locations -> States, it is linked to that State.\n2. If the value does not match an existing State, the State field is imported blank for the Lead - no error is raised.'],
    ['City*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. City must have a value.\n2. Value must not exceed 100 characters.\n3. Accepts free text; not validated against a master.'],
    ['Zip Code', 'Free Text', '1-20 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 20 characters.'],
    ['Address Line 1*', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Mandatory', '1. Address Line 1 must have a value.\n2. Value must not exceed 500 characters.'],
    ['Address Line 2', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 500 characters.'],
    ['Address Line 3', 'Free Text', '1-500 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 500 characters.'],
    ['Company Email', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If entered, Company Email must be a valid email format.\n2. Value must not exceed 100 characters.\n3. Duplicate Company Email across Leads is allowed (no uniqueness check).'],
    ['Company Contact No.', 'Free Text', '1-20 Numeric & Special Characters', 'Optional', '1. Value must not exceed 20 characters.\n2. Duplicate Company Contact No. across Leads is allowed (no uniqueness check).'],
    ['Company Website', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Optional', '1. If entered, Company Website must be a valid URL/website format.\n2. Value must not exceed 100 characters.'],
    ['Source', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Source options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Lead Source.\n2. One of the options should be selected from the dropdown.\n3. If no option is selected or an incorrect value is entered, Source shall be empty for the Lead (no blockage).'],
    ['Lead Owner Email Id*', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Mandatory', '1. Lead Owner Email Id must have a value.\n2. Value must be a valid email format.\n3. Email ID must match the email of an existing user. Source: FMS -> Admin -> Company Details -> Manage My Users.\n4. Value must not exceed 100 characters.'],
    ['Lead Stage*', 'Dropdown (Real-time, Lead Stage Master)', 'Value from Dropdown', 'Mandatory', '1. Lead Stage must have a value.\n2. Each cell of the column contains the Lead Stage options configured under Sales -> Settings -> Lead Stages.\n3. Value must match one of the configured Lead Stage options exactly.'],
    ['Industry', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Industry options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Industry.\n2. If no option is selected or an incorrect value is entered, Industry shall be empty for the Lead (no blockage).'],
    ['Lead Temperature', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Lead Temperature options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Opportunity Rating.\n2. If left blank or an incorrect value is entered, Lead Temperature defaults to \'Hot\'.'],
    ['Service Interest', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Service Interest options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Service Interest.\n2. If no option is selected or an incorrect value is entered, Service Interest shall be empty for the Lead (no blockage).'],
    ['Networks', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Each cell of the column contains the Networks options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Networks.\n2. If no option is selected or an incorrect value is entered, Networks shall be empty for the Lead (no blockage).'],
    ['Notes', 'Free Text', '1-1000 Alpha Numeric & Special Characters', 'Optional', '1. Value must not exceed 1000 characters.'],
    ['Primary Contact', 'Boolean Dropdown', 'TRUE / FALSE', 'Optional', '1. Only TRUE, FALSE, or blank is accepted.\n2. If TRUE: First Name, Last Name and Email become Mandatory for that row. If any of the three is left blank, a mandatory-field validation is triggered for the missing field(s).\n3. If FALSE or left blank, any values entered in First Name, Last Name, Email, Phone No. (Primary) and Designation are ignored and not saved against the Lead - these fields are only processed when Primary Contact = TRUE.'],
    ['First Name', 'Free Text', '1-50 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. Value must not exceed 50 characters.'],
    ['Last Name', 'Free Text', '1-50 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. Value must not exceed 50 characters.'],
    ['Email', 'Free Text', '1-100 Alpha Numeric & Special Characters', 'Conditionally Mandatory (Mandatory when Primary Contact = TRUE)', '1. Mandatory when Primary Contact = TRUE; otherwise Optional/ignored.\n2. If entered, Email must be a valid email format.\n3. Value must not exceed 100 characters.'],
    ['Phone No. (Primary)', 'Free Text', '1-20 Numeric & Special Characters', 'Optional', '1. Only processed when Primary Contact = TRUE (otherwise ignored - see Primary Contact validations).\n2. Phone Number must contain 20 or fewer characters.\n3. Must contain digits and permitted special characters only (+, spaces, hyphens, parentheses) - no alphabets.'],
    ['Designation', 'Dropdown (Real-time, Company Master)', 'Value from Dropdown', 'Optional', '1. Only processed when Primary Contact = TRUE (otherwise ignored).\n2. Each cell of the column contains the Designation options configured for the company. Source: FMS -> Master -> Others -> User Value Detail -> Designation.\n3. If no option is selected or an incorrect value is entered, Designation shall be empty for the Lead (no blockage).']
  ];

  const ws2 = XLSX.utils.aoa_to_sheet([instructionHeaders, ...instructionRows]);
  ws2['!cols'] = [
    { wch: 22 },
    { wch: 25 },
    { wch: 35 },
    { wch: 28 },
    { wch: 70 }
  ];

  XLSX.utils.book_append_sheet(wb, ws1, 'Import Bulk Lead');
  XLSX.utils.book_append_sheet(wb, ws2, 'Instructions');

  XLSX.writeFile(wb, 'Lead_Bulk_Import_Template.xlsx');
  showToast('Template Downloaded', 'Lead_Bulk_Import_Template.xlsx with 2 sheets generated.', 'success');
}

// Quick Sample Loaders for Testing Scenarios directly in the browser
window.loadSampleScenario = function(type) {
  if (!importPermissionEnabled) {
    showToast(
      'Access Restricted',
      'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance.',
      'error'
    );
    return;
  }
  openDrawer();

  let sampleData = [];
  let fileName = 'Sample_Leads.xlsx';

  if (type === 'valid') {
    fileName = 'Valid_3_Leads.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        'Nexus Global Freight Ltd', 'NGF', 'Nexus Global Freight Limited', 'Customer', 'US', 'California', 'Los Angeles', '90001', '500 Harbor Dr', '', '', 'info@nexusfreight.com', '+1 310-555-1234', 'www.nexusfreight.com', 'Website', 'sakshi.barnwal@freightoscope.com', 'New', 'Freight Forwarding', 'Hot', 'Air Freight', 'IATA', 'Key account prospect', 'TRUE', 'Sharon', 'Roy', 'sharon.roy@nexusfreight.com', '+1 310-555-9876', 'Logistics Director'
      ],
      [
        'Blue Ocean Shipping LLC', 'BOS', 'Blue Ocean Shipping LLC', 'Customer', 'AE', 'Dubai', 'Dubai', '00000', 'Office 402, Business Bay Tower', '', '', 'ops@blueocean.ae', '+971 4 555 1234', 'www.blueocean.ae', 'Referral', 'rafia.khan@freightoscope.com', 'Qualified', 'Shipping', 'Hot', 'Ocean Freight (FCL)', 'FIATA', 'Ready for quote', 'FALSE', '', '', '', '', ''
      ],
      [
        'Apex World Transports Pvt', 'AWT', 'Apex World Transports Private Limited', 'Agent', 'IN', 'Maharashtra', 'Mumbai', '400001', '12 Nariman Point, Express Towers', '', '', 'contact@apexworld.in', '+91 22 5555 1234', 'www.apexworld.in', 'Sales Call', 'john.doe@freightoscope.com', 'New', 'Logistics', 'Warm', 'Customs Clearance', 'CII', 'Expansion in India ports', 'TRUE', 'Rajesh', 'Kumar', 'rajesh.k@apexworld.in', '+91 98200 12345', 'Managing Partner'
      ]
    ];
  } else if (type === 'errors') {
    fileName = 'Sample_With_Field_Errors.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        '', // Missing Company Name
        'NGF', 'Nexus Global', 'Customer',
        'USA', // Invalid 2-letter Country Code
        'CA', 'Los Angeles', '90001',
        '', // Missing Address Line 1
        '', '',
        'not-an-email', // Invalid Email format
        '',
        'http://nexus.com', '',
        'unknown.user@freightoscope.com', // Owner Not Found
        '', // Missing Lead Stage
        '', '', '', '', '',
        'TRUE',
        '', // Missing First Name (Primary Contact = TRUE)
        'Roy',
        'sharon-email', // Invalid Contact Email
        '+1 310-CALL-NOW', // Letters in phone number
        'Director'
      ],
      [
        'Global Express Cargo Corporation International Overseas Worldwide Global Line Limited Company Extra Name Exceeding Limit Very Long Name Over Hundred Characters Completely', // > 100 chars
        'GEC', 'Global Express', 'Customer',
        'ZZ', // Country Not Found
        'NY',
        'New York City Central Metropolitan Regional Area Exceeding The One Hundred Character Limit Length Test Checking Rule', // City > 100 chars
        '10001',
        '100 Broadway St', '', '',
        'info@globalexpress.com', '', 'www.globalexpress.com', '',
        'invalid-email-format', // Invalid Owner Email format
        'New', '', '', '', '', '',
        'MAYBE', // Invalid Primary Contact boolean value
        '', '', '', '', ''
      ]
    ];
  } else if (type === 'duplicates') {
    fileName = 'Sample_With_Duplicate_Leads.xlsx';
    sampleData = [
      DEFINED_HEADERS,
      [
        'Evergreen Freight Forwarders', 'EFF', 'Evergreen Freight Forwarders', 'Customer', 'IN', 'Tamilnadu', 'Chennai', '600001', '14 Rajaji Salai', '', '', 'info@evergreen.in', '', 'www.evergreen.in', '', 'sakshi.barnwal@freightoscope.com', 'New', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ],
      [
        'Pacific Trans Logistics', 'PTL', 'Pacific Trans', 'Customer', 'US', 'WA', 'Seattle', '98101', '400 Pine St', '', '', 'info@pacific.com', '', 'www.pacific.com', '', 'rafia.khan@freightoscope.com', 'Qualified', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ],
      [
        'Evergreen Freight Forwarders ', 'EFF', 'Evergreen Freight Forwarders', 'Customer', 'IN', 'Tamilnadu', ' Chennai ', '600001', '14 Rajaji Salai', '', '', 'contact@evergreen.in', '', 'www.evergreen.in', '', 'sakshi.barnwal@freightoscope.com', 'New', '', 'Hot', '', '', '', 'FALSE', '', '', '', '', ''
      ]
    ];
  } else if (type === 'empty') {
    fileName = 'Empty_Lead_Template.xlsx';
    sampleData = [
      DEFINED_HEADERS
      // No rows
    ];
  } else if (type === 'wrong_headers') {
    fileName = 'Corrupted_Headers.xlsx';
    sampleData = [
      ['Customer Name', 'City Location', 'Phone', 'Email Address'],
      ['Acme Corp', 'New York', '1234567', 'test@acme.com']
    ];
  }

  // Create virtual file and trigger parsing
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  XLSX.utils.book_append_sheet(wb, ws, 'Import Bulk Lead');

  // UI simulate file
  uploadedFileName.textContent = fileName;
  uploadedFileSize.textContent = '14.8 KB';
  uploadedFileBar.style.display = 'flex';
  btnDrawerReset.style.display = 'inline-block';
  hideBanners();
  hideErrors();
  hideSuccessBox();

  processWorkbook(wb);
};

// Toast Notifications System
function showToast(title, message, type = 'info', duration = 4500) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  let icon = 'ℹ️';
  if (type === 'success') icon = '✅';
  if (type === 'error') icon = '❌';
  if (type === 'warning') icon = '⚠️';

  toast.innerHTML = `
    <div class="toast-icon">${icon}</div>
    <div class="toast-content">
      <div class="toast-title">${escapeHtml(title)}</div>
      <div class="toast-desc">${escapeHtml(message)}</div>
    </div>
    <button class="toast-close">&times;</button>
  `;

  toast.querySelector('.toast-close').addEventListener('click', () => {
    toast.remove();
  });

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* ==========================================================================
   FREIGHTOSCOPE OPPORTUNITIES PIPELINE & TABLE VIEW PROTOTYPE ENGINE
   (Matches Screenshot 1 & 2 Fidelity, Drag & Drop, Permissions & JT Flow)
   ========================================================================== */

// Configured Opportunity Stages (Ordered Left -> Right, Active Only per Requirement 3)
let opportunityStages = [
  { id: 'new', name: 'New', color: '#475467', order: 1, active: true },
  { id: 'qualification', name: 'Qualification', color: '#2563eb', order: 2, active: true },
  { id: 'quotation_sent', name: 'Quotation Sent', color: '#1d4ed8', order: 3, active: true },
  { id: 'won', name: 'Won', color: '#16a34a', order: 4, active: true },
  { id: 'lost', name: 'Lost', color: '#dc2626', order: 5, active: true }
];

// Initial Dummy Opportunities in Pipeline (Exact match to Screenshot 1 + Won stage)
const INITIAL_OPPORTUNITIES = [
  // Stage: New (2 Opportunities - USD 78,15,300)
  {
    id: 'OPP-101',
    name: 'ABC Electronics – Air Export',
    fullName: 'ABC Electronics – Air Export',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 4500000,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'new',
    createdDate: '2026-09-12T10:30:00',
    auditHistory: [
      { timestamp: '12-Sep-2026 10:30', text: 'Opportunity created in New stage by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-102',
    name: 'RGB Exports',
    fullName: 'RGB Exports Global Air & Ocean',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 3315300,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'new',
    createdDate: '2026-09-11T09:15:00',
    auditHistory: [
      { timestamp: '11-Sep-2026 09:15', text: 'Opportunity created in New stage by Sakshi Barnwal' }
    ]
  },

  // Stage: Qualification (3 Opportunities - USD 28,29,847)
  {
    id: 'OPP-201',
    name: 'ABC Electronics – Air Export',
    fullName: 'ABC Electronics – Air Export Phase 2',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 1000000,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'qualification',
    createdDate: '2026-09-12T11:45:00',
    auditHistory: [
      { timestamp: '12-Sep-2026 11:45', text: 'Stage moved to Qualification by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-202',
    name: 'ABC Electronics – Air Export',
    fullName: 'ABC Electronics – Air Export Commercial',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 929847,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'qualification',
    createdDate: '2026-09-10T14:20:00',
    auditHistory: [
      { timestamp: '10-Sep-2026 14:20', text: 'Opportunity created by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-203',
    name: 'Apex World Logistics - FCL Ocean',
    fullName: 'Apex World Logistics - FCL Ocean Freight',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 900000,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'qualification',
    createdDate: '2026-09-08T16:00:00',
    auditHistory: [
      { timestamp: '08-Sep-2026 16:00', text: 'Opportunity created by Sakshi Barnwal' }
    ]
  },

  // Stage: Quotation Sent (3 Opportunities - USD 27,71,747)
  {
    id: 'OPP-301',
    name: 'Dubai to Bangalore Door-to-D...',
    fullName: 'Dubai to Bangalore Door-to-Door Cargo',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 1200000,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'quotation_sent',
    createdDate: '2026-09-12T12:00:00',
    auditHistory: [
      { timestamp: '12-Sep-2026 12:00', text: 'Quotation sent to party by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-302',
    name: 'Delhi to London Air Export',
    fullName: 'Delhi to London Air Export Cargo',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 871747,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'quotation_sent',
    createdDate: '2026-09-09T15:30:00',
    auditHistory: [
      { timestamp: '09-Sep-2026 15:30', text: 'Quotation dispatched by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-303',
    name: 'Frankfurt to Mumbai FCL',
    fullName: 'Frankfurt to Mumbai FCL Shipment',
    party: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 700000,
    expProfitDisplay: 'USD 2,78,11,45,800',
    expProfitValue: 2781145800,
    closureDate: '31-Aug-2026',
    stage: 'quotation_sent',
    createdDate: '2026-09-07T11:10:00',
    auditHistory: [
      { timestamp: '07-Sep-2026 11:10', text: 'Quotation generated by Sakshi Barnwal' }
    ]
  },

  // Stage: Won (2 Opportunities - USD 45,60,000 - Added per User Request)
  {
    id: 'OPP-401',
    name: 'Pacific Trans Logistics - Sea Freight',
    fullName: 'Pacific Trans Logistics - Pacific Sea Freight Contract',
    party: 'Pacific Trans Logistics Inc, Seattle, USA',
    tradeLane: 'USA - Japan',
    expRevDisplay: 'USD 25,50,000',
    expRevValue: 2550000,
    expProfitDisplay: 'USD 6,80,000',
    expProfitValue: 680000,
    closureDate: '28-Aug-2026',
    stage: 'won',
    createdDate: '2026-09-11T14:00:00',
    auditHistory: [
      { timestamp: '11-Sep-2026 14:00', text: 'Opportunity marked WON by Sakshi Barnwal' }
    ]
  },
  {
    id: 'OPP-402',
    name: 'Evergreen Global Cargo - Air Import',
    fullName: 'Evergreen Global Cargo - Air Import Express',
    party: 'Evergreen Cargo Co, Hamburg, Germany',
    tradeLane: 'Germany - Singapore',
    expRevDisplay: 'USD 20,10,000',
    expRevValue: 2010000,
    expProfitDisplay: 'USD 4,75,000',
    expProfitValue: 475000,
    closureDate: '25-Aug-2026',
    stage: 'won',
    createdDate: '2026-09-06T10:00:00',
    auditHistory: [
      { timestamp: '06-Sep-2026 10:00', text: 'Contract signed, marked WON by Sakshi Barnwal' }
    ]
  },

  // Stage: Lost (1 Opportunity - USD 1,78,11,56,800)
  {
    id: 'OPP-501',
    name: 'Mumbai to Singap...',
    fullName: 'Mumbai to Singapore Door-to-Door Air Cargo',
    party: 'RGB Exports Flo...',
    fullNameParty: 'RGB Exports Florida, USA',
    tradeLane: 'India - Germany',
    expRevDisplay: 'USD 1,78,11,56,800',
    expRevValue: 1781156800,
    expProfitDisplay: '--',
    expProfitValue: 0,
    closureDate: '31-Aug-2026',
    stage: 'lost',
    createdDate: '2026-09-05T08:00:00',
    auditHistory: [
      { timestamp: '05-Sep-2026 08:00', text: 'Opportunity marked as Lost by Sakshi Barnwal' }
    ]
  }
];

// Opportunities Prototype State
let opportunitiesData = JSON.parse(JSON.stringify(INITIAL_OPPORTUNITIES));
let oppUpdatePermissionEnabled = true; // Controlled by Admin -> Departments & Roles -> Sales Opportunity -> Update
let oppUpdatePermissionDraft = true;
let activeOpportunityView = 'pipeline'; // 'pipeline' or 'table'
let currentSearchQuery = '';
let draggedCardOppId = null;
let activeMenuOppId = null;

// Currency & Formatting Helpers (Indian/Western Commas matching Screenshot 1)
function formatIndianCurrency(num, currency = 'USD') {
  if (num === null || num === undefined || isNaN(num)) return '--';
  const n = Math.round(num).toString();
  let lastThree = n.substring(n.length - 3);
  const otherNumbers = n.substring(0, n.length - 3);
  if (otherNumbers !== '') {
    lastThree = ',' + lastThree;
  }
  const formatted = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + lastThree;
  return `${currency} ${formatted}`;
}

// Initialize Opportunities Module
function initOpportunitiesModule() {
  // Load saved view from localStorage (Requirement 10: Persistence)
  const savedView = localStorage.getItem('fos_preferred_opportunity_view');
  if (savedView && (savedView === 'pipeline' || savedView === 'table')) {
    activeOpportunityView = savedView;
  } else {
    // Screenshot 1 displays Pipeline View as active
    activeOpportunityView = 'pipeline';
  }

  // Bind Sidebar Navigation
  setupSidebarNavigation();

  // Bind View Switchers (Requirement 1 & 2)
  const btnViewTable = document.getElementById('btnViewTable');
  const btnViewPipeline = document.getElementById('btnViewPipeline');
  if (btnViewTable) {
    btnViewTable.addEventListener('click', () => switchOpportunityView('table'));
  }
  if (btnViewPipeline) {
    btnViewPipeline.addEventListener('click', () => switchOpportunityView('pipeline'));
  }

  // Bind Search Input
  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      if (activeModule === 'contacts') {
        contactSearchQuery = q;
        renderContactsTable();
      } else if (activeModule === 'leads') {
        filterLeadsTable(q);
      } else {
        currentSearchQuery = q;
        renderCurrentOpportunityView();
      }
    });
  }

  // Bind Scroll Arrows (Requirement 3)
  const scrollArrowLeft = document.getElementById('scrollArrowLeft');
  const scrollArrowRight = document.getElementById('scrollArrowRight');
  const pipelineScrollContainer = document.getElementById('pipelineScrollContainer');
  if (scrollArrowLeft && pipelineScrollContainer) {
    scrollArrowLeft.addEventListener('click', () => {
      pipelineScrollContainer.scrollBy({ left: -300, behavior: 'smooth' });
    });
  }
  if (scrollArrowRight && pipelineScrollContainer) {
    scrollArrowRight.addEventListener('click', () => {
      pipelineScrollContainer.scrollBy({ left: 300, behavior: 'smooth' });
    });
  }
  if (pipelineScrollContainer) {
    pipelineScrollContainer.addEventListener('scroll', updateScrollArrows);
  }

  // Bind Create Opportunity Modal
  const btnCreateOpp = document.getElementById('btnCreateOpp');
  const createOppModal = document.getElementById('createOppModal');
  const btnCloseCreateOpp = document.getElementById('btnCloseCreateOpp');
  const btnCancelCreateOpp = document.getElementById('btnCancelCreateOpp');
  const btnSubmitCreateOpp = document.getElementById('btnSubmitCreateOpp');

  if (btnCreateOpp) {
    btnCreateOpp.addEventListener('click', () => {
      if (createOppModal) createOppModal.classList.add('open');
      const backdrop = document.getElementById('oppDetailsBackdrop');
      if (backdrop) backdrop.style.display = 'block';
    });
  }
  if (btnCloseCreateOpp) btnCloseCreateOpp.addEventListener('click', closeCreateOppModal);
  if (btnCancelCreateOpp) btnCancelCreateOpp.addEventListener('click', closeCreateOppModal);
  if (btnSubmitCreateOpp) btnSubmitCreateOpp.addEventListener('click', handleCreateOppSubmit);

  // Bind Opportunity Details Modal Close
  const btnCloseOppDetails = document.getElementById('btnCloseOppDetails');
  const btnFooterCloseOppDetails = document.getElementById('btnFooterCloseOppDetails');
  const oppDetailsBackdrop = document.getElementById('oppDetailsBackdrop');
  if (btnCloseOppDetails) btnCloseOppDetails.addEventListener('click', closeOppDetailsModal);
  if (btnFooterCloseOppDetails) btnFooterCloseOppDetails.addEventListener('click', closeOppDetailsModal);
  if (oppDetailsBackdrop) {
    oppDetailsBackdrop.addEventListener('click', () => {
      closeOppDetailsModal();
      closeCreateOppModal();
    });
  }

  // Bind No Active Stages Empty-State Ok Button (Requirement 9)
  const btnAcknowledgeNoStages = document.getElementById('btnAcknowledgeNoStages');
  if (btnAcknowledgeNoStages) {
    btnAcknowledgeNoStages.addEventListener('click', () => {
      showToast('Acknowledged', 'You remain in Pipeline View. You can switch back to Table View at any time.', 'info', 3000);
    });
  }

  // Close menus on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.card-menu-btn') && !e.target.closest('.card-menu-popover')) {
      closeAllOppMenus();
    }
  });

  // Initial Render
  renderCurrentOpportunityView();
  updateOppPermissionPillUI();
}

// Sidebar Navigation Handling
function setupSidebarNavigation() {
  const subitemContacts = document.getElementById('subitemContacts');
  const subitemOpportunities = document.getElementById('subitemOpportunities');
  const subitemLeads = document.getElementById('subitemLeads');

  if (subitemContacts) {
    subitemContacts.addEventListener('click', () => {
      showContactsModule();
    });
  }

  if (subitemOpportunities) {
    subitemOpportunities.addEventListener('click', () => {
      showOpportunitiesModule();
    });
  }

  if (subitemLeads) {
    subitemLeads.addEventListener('click', () => {
      showLeadsModule();
    });
  }
}

function showContactsModule() {
  activeModule = 'contacts';
  const subitemContacts = document.getElementById('subitemContacts');
  const subitemOpportunities = document.getElementById('subitemOpportunities');
  const subitemLeads = document.getElementById('subitemLeads');
  const pageTitle = document.getElementById('pageTitle');
  const contactsHeaderActions = document.getElementById('contactsHeaderActions');
  const oppHeaderActions = document.getElementById('oppHeaderActions');
  const leadsHeaderActions = document.getElementById('leadsHeaderActions');
  const viewSwitcherGroup = document.getElementById('viewSwitcherGroup');
  const pipelineView = document.getElementById('pipelineView');
  const oppTableView = document.getElementById('oppTableView');
  const leadsTableView = document.getElementById('leadsTableView');
  const contactsTableView = document.getElementById('contactsTableView');
  const searchInput = document.getElementById('searchInput');
  const dateRangeText = document.getElementById('dateRangeText');

  if (subitemContacts) subitemContacts.classList.add('active');
  if (subitemOpportunities) subitemOpportunities.classList.remove('active');
  if (subitemLeads) subitemLeads.classList.remove('active');
  if (pageTitle) pageTitle.textContent = 'Contacts';
  if (contactsHeaderActions) contactsHeaderActions.style.display = 'flex';
  if (oppHeaderActions) oppHeaderActions.style.display = 'none';
  if (leadsHeaderActions) leadsHeaderActions.style.display = 'none';
  if (viewSwitcherGroup) viewSwitcherGroup.style.display = 'none';
  if (pipelineView) pipelineView.style.display = 'none';
  if (oppTableView) oppTableView.style.display = 'none';
  if (leadsTableView) leadsTableView.style.display = 'none';
  if (contactsTableView) contactsTableView.style.display = 'block';
  if (searchInput) {
    searchInput.placeholder = 'Search Contacts';
    searchInput.value = contactSearchQuery;
  }
  if (dateRangeText) dateRangeText.textContent = '9/1/2026 - 9/30/2026';

  renderContactsTable();
}

function showOpportunitiesModule() {
  activeModule = 'opportunities';
  const subitemContacts = document.getElementById('subitemContacts');
  const subitemOpportunities = document.getElementById('subitemOpportunities');
  const subitemLeads = document.getElementById('subitemLeads');
  const pageTitle = document.getElementById('pageTitle');
  const contactsHeaderActions = document.getElementById('contactsHeaderActions');
  const oppHeaderActions = document.getElementById('oppHeaderActions');
  const leadsHeaderActions = document.getElementById('leadsHeaderActions');
  const viewSwitcherGroup = document.getElementById('viewSwitcherGroup');
  const leadsTableView = document.getElementById('leadsTableView');
  const contactsTableView = document.getElementById('contactsTableView');
  const searchInput = document.getElementById('searchInput');
  const dateRangeText = document.getElementById('dateRangeText');

  if (subitemContacts) subitemContacts.classList.remove('active');
  if (subitemOpportunities) subitemOpportunities.classList.add('active');
  if (subitemLeads) subitemLeads.classList.remove('active');
  if (pageTitle) pageTitle.textContent = 'Opportunities';
  if (contactsHeaderActions) contactsHeaderActions.style.display = 'none';
  if (oppHeaderActions) oppHeaderActions.style.display = 'flex';
  if (leadsHeaderActions) leadsHeaderActions.style.display = 'none';
  if (viewSwitcherGroup) viewSwitcherGroup.style.display = 'inline-flex';
  if (leadsTableView) leadsTableView.style.display = 'none';
  if (contactsTableView) contactsTableView.style.display = 'none';
  if (searchInput) {
    searchInput.placeholder = 'Search Opportunities';
    searchInput.value = currentSearchQuery;
  }
  if (dateRangeText) dateRangeText.textContent = '9/11/2026 - 9/12/2026';

  renderCurrentOpportunityView();
}

function showLeadsModule() {
  activeModule = 'leads';
  const subitemContacts = document.getElementById('subitemContacts');
  const subitemOpportunities = document.getElementById('subitemOpportunities');
  const subitemLeads = document.getElementById('subitemLeads');
  const pageTitle = document.getElementById('pageTitle');
  const contactsHeaderActions = document.getElementById('contactsHeaderActions');
  const oppHeaderActions = document.getElementById('oppHeaderActions');
  const leadsHeaderActions = document.getElementById('leadsHeaderActions');
  const viewSwitcherGroup = document.getElementById('viewSwitcherGroup');
  const pipelineView = document.getElementById('pipelineView');
  const oppTableView = document.getElementById('oppTableView');
  const leadsTableView = document.getElementById('leadsTableView');
  const contactsTableView = document.getElementById('contactsTableView');
  const searchInput = document.getElementById('searchInput');
  const dateRangeText = document.getElementById('dateRangeText');

  if (subitemContacts) subitemContacts.classList.remove('active');
  if (subitemOpportunities) subitemOpportunities.classList.remove('active');
  if (subitemLeads) subitemLeads.classList.add('active');
  if (pageTitle) pageTitle.textContent = 'Leads';
  if (contactsHeaderActions) contactsHeaderActions.style.display = 'none';
  if (oppHeaderActions) oppHeaderActions.style.display = 'none';
  if (leadsHeaderActions) leadsHeaderActions.style.display = 'flex';
  if (viewSwitcherGroup) viewSwitcherGroup.style.display = 'none';
  if (pipelineView) pipelineView.style.display = 'none';
  if (oppTableView) oppTableView.style.display = 'none';
  if (contactsTableView) contactsTableView.style.display = 'none';
  if (leadsTableView) leadsTableView.style.display = 'block';
  if (searchInput) searchInput.placeholder = 'Search Leads';
  if (dateRangeText) dateRangeText.textContent = '9/11/2026 - 9/12/2026';

  renderLeadsTable();
}

// Switch Opportunity View (Requirement 2 & 10)
function switchOpportunityView(newView) {
  if (newView === activeOpportunityView) return;

  activeOpportunityView = newView;
  localStorage.setItem('fos_preferred_opportunity_view', newView);

  // Clear search conditions upon switching views (Requirement 2)
  currentSearchQuery = '';
  const searchInput = document.getElementById('searchInput');
  if (searchInput) searchInput.value = '';

  renderCurrentOpportunityView();
}

// Render Current Opportunity View
function renderCurrentOpportunityView() {
  const btnViewTable = document.getElementById('btnViewTable');
  const btnViewPipeline = document.getElementById('btnViewPipeline');
  const pipelineView = document.getElementById('pipelineView');
  const oppTableView = document.getElementById('oppTableView');

  if (activeOpportunityView === 'pipeline') {
    if (btnViewPipeline) btnViewPipeline.classList.add('active');
    if (btnViewTable) btnViewTable.classList.remove('active');
    if (pipelineView) pipelineView.style.display = 'flex';
    if (oppTableView) oppTableView.style.display = 'none';
    renderPipelineView();
  } else {
    if (btnViewTable) btnViewTable.classList.add('active');
    if (btnViewPipeline) btnViewPipeline.classList.remove('active');
    if (pipelineView) pipelineView.style.display = 'none';
    if (oppTableView) oppTableView.style.display = 'block';
    renderOppTableView();
  }
}

// Calculate Stage Total Revenue (Requirement 3: sum of opportunities converted to branch currency)
function calculateStageTotal(stageId) {
  const stageOpps = opportunitiesData.filter(o => o.stage === stageId);
  let total = 0;
  stageOpps.forEach(opp => {
    total += (opp.expRevValue || 0);
  });
  return total;
}

// Render Pipeline View (Requirement 3 & 4)
function renderPipelineView() {
  const container = document.getElementById('pipelineScrollContainer');
  const noActiveStagesBox = document.getElementById('noActiveStagesContainer');
  const scrollArrowLeft = document.getElementById('scrollArrowLeft');
  const scrollArrowRight = document.getElementById('scrollArrowRight');

  if (!container) return;

  // Check Active Stages (Requirement 9: No Active Opportunity Stages)
  const activeStages = opportunityStages
    .filter(s => s.active)
    .sort((a, b) => a.order - b.order);

  if (activeStages.length === 0) {
    container.style.display = 'none';
    if (noActiveStagesBox) noActiveStagesBox.style.display = 'flex';
    if (scrollArrowLeft) scrollArrowLeft.style.display = 'none';
    if (scrollArrowRight) scrollArrowRight.style.display = 'none';
    return;
  }

  container.style.display = 'flex';
  if (noActiveStagesBox) noActiveStagesBox.style.display = 'none';
  if (scrollArrowLeft) scrollArrowLeft.style.display = 'flex';
  if (scrollArrowRight) scrollArrowRight.style.display = 'flex';

  container.innerHTML = '';

  // Render each active stage column from left to right
  activeStages.forEach(stage => {
    // Filter matching opportunities
    let opps = opportunitiesData.filter(o => o.stage === stage.id);

    // Apply Search Filter if present
    if (currentSearchQuery) {
      opps = opps.filter(o =>
        o.name.toLowerCase().includes(currentSearchQuery) ||
        o.party.toLowerCase().includes(currentSearchQuery) ||
        o.tradeLane.toLowerCase().includes(currentSearchQuery)
      );
    }

    // Order: Newest Created Date -> Oldest Created Date (Requirement 4)
    opps.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));

    // Calculate Stage Revenue
    const totalRev = calculateStageTotal(stage.id);
    const formattedTotal = formatIndianCurrency(totalRev);

    // Create Column Element
    const colEl = document.createElement('div');
    colEl.className = 'pipeline-stage-column';
    colEl.setAttribute('data-stage-id', stage.id);

    colEl.innerHTML = `
      <!-- Sticky Stage Header (Requirement 3 & Screenshot 1) -->
      <div class="stage-sticky-header">
        <div class="stage-header-title-row">
          <span class="stage-color-dot" style="background-color: ${stage.color};"></span>
          <span class="stage-title-text">${escapeHtml(stage.name)}</span>
        </div>
        <div class="stage-header-summary-row">
          <span>${opps.length} Opportunities - ${formattedTotal}</span>
        </div>
      </div>

      <!-- Stage Cards Drop Area -->
      <div class="stage-cards-list" id="stageList_${stage.id}" data-stage-id="${stage.id}">
        <!-- Cards injected dynamically -->
      </div>
    `;

    const cardsList = colEl.querySelector('.stage-cards-list');

    // Drag Over & Drop Events on Column Dropzone (Requirement 7)
    cardsList.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      cardsList.classList.add('drag-over');
    });

    cardsList.addEventListener('dragleave', (e) => {
      if (!cardsList.contains(e.relatedTarget)) {
        cardsList.classList.remove('drag-over');
      }
    });

    cardsList.addEventListener('drop', (e) => {
      e.preventDefault();
      cardsList.classList.remove('drag-over');
      handleCardDrop(stage.id);
    });

    // Render cards inside stage
    opps.forEach(opp => {
      const cardEl = createOpportunityCardElement(opp, stage);
      cardsList.appendChild(cardEl);
    });

    container.appendChild(colEl);
  });

  updateScrollArrows();
}

// Create Opportunity Card Element (Requirement 4, 5, 7, 8)
function createOpportunityCardElement(opp, stage) {
  const card = document.createElement('div');
  card.className = 'opportunity-card';
  card.id = `card_${opp.id}`;
  card.setAttribute('draggable', 'true');
  card.setAttribute('data-opp-id', opp.id);
  card.setAttribute('data-stage-id', stage.id);
  card.style.setProperty('--card-stage-color', stage.color);

  // Profit formatting: green if value exists, -- otherwise
  const profitHtml = opp.expProfitDisplay && opp.expProfitDisplay !== '--'
    ? `<span class="financial-value profit">${escapeHtml(opp.expProfitDisplay)}</span>`
    : `<span class="financial-value" style="color: #9ca3af;">--</span>`;

  card.innerHTML = `
    <!-- Top Row: Name & Three-Dot Menu -->
    <div class="card-header-row">
      <a class="card-title-link" title="${escapeHtml(opp.fullName || opp.name)}">${escapeHtml(opp.name)}</a>
      <button type="button" class="card-menu-btn" title="More Actions">⋮</button>
      
      <!-- Three-Dot Menu Popover (Screenshot 1 Match) -->
      <div class="card-menu-popover" id="menu_${opp.id}">
        <div class="menu-popover-item" data-action="edit">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
          </svg>
          <span>Edit</span>
        </div>
        <div class="menu-popover-item" data-action="view">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
            <circle cx="12" cy="12" r="3"></circle>
          </svg>
          <span>View</span>
        </div>
        <div class="menu-popover-item" data-action="duplicate">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
          </svg>
          <span>Duplicate</span>
        </div>
        <div class="menu-popover-item danger" data-action="delete">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
          <span>Delete</span>
        </div>
        <div class="menu-popover-item" data-action="inquiry">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
          </svg>
          <span>Create Inquiry</span>
        </div>
        <div class="menu-popover-item" data-action="quotation">
          <svg class="menu-popover-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
          </svg>
          <span>Create Quotation</span>
        </div>
      </div>
    </div>

    <!-- Row 2: Party Name (Briefcase Icon) -->
    <div class="card-meta-row" title="${escapeHtml(opp.fullNameParty || opp.party)}">
      <svg class="card-meta-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect>
        <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>
      </svg>
      <span class="card-meta-text">${escapeHtml(opp.party)}</span>
    </div>

    <!-- Row 3: Trade Lane (Globe Icon) -->
    <div class="card-meta-row" title="${escapeHtml(opp.tradeLane)}">
      <svg class="card-meta-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10"></circle>
        <line x1="2" y1="12" x2="22" y2="12"></line>
        <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>
      </svg>
      <span class="card-meta-text">${escapeHtml(opp.tradeLane)}</span>
    </div>

    <!-- Row 4: Financials (Expected Revenue & Expected Profit) -->
    <div class="card-financials-grid">
      <div class="financial-col">
        <span class="financial-label">Expected Revenue</span>
        <span class="financial-value">${escapeHtml(opp.expRevDisplay)}</span>
      </div>
      <div class="financial-col">
        <span class="financial-label">Expected Profit</span>
        ${profitHtml}
      </div>
    </div>

    <!-- Row 5: Estimated Closure Date -->
    <div class="card-date-row">
      <span class="card-date-label">Estimated Closure Date</span>
      <span class="card-date-value">${escapeHtml(opp.closureDate || '--')}</span>
    </div>
  `;

  // Opportunity Name Click -> Open Details (Requirement 4: Does not trigger drag)
  const titleLink = card.querySelector('.card-title-link');
  titleLink.addEventListener('click', (e) => {
    e.stopPropagation();
    openOppDetailsModal(opp.id);
  });

  // Three-dot Menu Click -> Open Menu (Requirement 7: Does not trigger drag)
  const menuBtn = card.querySelector('.card-menu-btn');
  const menuPopover = card.querySelector('.card-menu-popover');
  menuBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const isAlreadyOpen = menuPopover.classList.contains('open');
    closeAllOppMenus();
    if (!isAlreadyOpen) {
      menuPopover.classList.add('open');
      activeMenuOppId = opp.id;
    }
  });

  // Menu Items Action Clicks
  card.querySelectorAll('.menu-popover-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.stopPropagation();
      const action = item.getAttribute('data-action');
      closeAllOppMenus();
      handleOppAction(action, opp.id);
    });
  });

  // Drag & Drop Handlers (Requirement 7 & 8)
  card.addEventListener('dragstart', (e) => {
    // If dragging initiated from title or menu button, prevent drag
    if (e.target.closest('.card-title-link') || e.target.closest('.card-menu-btn') || e.target.closest('.card-menu-popover')) {
      e.preventDefault();
      return;
    }

    draggedCardOppId = opp.id;
    card.classList.add('dragging');
    e.dataTransfer.setData('text/plain', opp.id);
    e.dataTransfer.effectAllowed = 'move';
  });

  card.addEventListener('dragend', () => {
    card.classList.remove('dragging');
    draggedCardOppId = null;
    document.querySelectorAll('.stage-cards-list').forEach(l => l.classList.remove('drag-over'));
  });

  return card;
}

// Close all popover menus
function closeAllOppMenus() {
  document.querySelectorAll('.card-menu-popover.open').forEach(m => m.classList.remove('open'));
  activeMenuOppId = null;
}

// Handle Card Drop (Requirement 7 & 8: Successful Update vs Permission Denied)
function handleCardDrop(targetStageId) {
  if (!draggedCardOppId) return;

  const oppIndex = opportunitiesData.findIndex(o => o.id === draggedCardOppId);
  if (oppIndex === -1) return;

  const opp = opportunitiesData[oppIndex];
  const sourceStageId = opp.stage;

  // Dropping into the same stage -> no-op
  if (sourceStageId === targetStageId) return;

  // Check Stage Update Permission (Requirement 8)
  if (!oppUpdatePermissionEnabled) {
    // Permission Denied Flow:
    // 1. Card returns to original position immediately
    // 2. Shake animation
    // 3. Show tooltip: "You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance."
    const originalCard = document.getElementById(`card_${opp.id}`);
    if (originalCard) {
      originalCard.classList.add('snap-back');
      setTimeout(() => originalCard.classList.remove('snap-back'), 450);

      // Render Permission Tooltip over the card
      showCardPermissionTooltip(originalCard);
    }

    showToast(
      'Permission Denied',
      'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance.',
      'error',
      5500
    );
    return;
  }

  // Permission Granted -> Update Stage immediately (Requirement 7)
  const sourceStage = opportunityStages.find(s => s.id === sourceStageId);
  const targetStage = opportunityStages.find(s => s.id === targetStageId);

  opp.stage = targetStageId;

  // Record change in audit history (Requirement 7)
  const now = new Date();
  const timeStr = `${now.getDate()}-${now.toLocaleString('default', { month: 'short' })}-${now.getFullYear()} ${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')}`;
  if (!opp.auditHistory) opp.auditHistory = [];
  opp.auditHistory.unshift({
    timestamp: timeStr,
    text: `Stage moved from ${sourceStage ? sourceStage.name : sourceStageId} to ${targetStage ? targetStage.name : targetStageId} by Sakshi Barnwal`
  });

  // Re-render pipeline view immediately
  renderPipelineView();

  showToast(
    'Stage Updated',
    `"${opp.name}" moved to ${targetStage ? targetStage.name : targetStageId}`,
    'success',
    3500
  );
}

// Display Permission Denied Tooltip on Card (Requirement 8)
function showCardPermissionTooltip(cardEl) {
  // Remove any existing card tooltips
  document.querySelectorAll('.permission-denied-tooltip').forEach(t => t.remove());

  const tooltip = document.createElement('div');
  tooltip.className = 'permission-denied-tooltip';
  tooltip.textContent = 'You do not have permission to perform this action. Please contact your Admin or service@freightoscope.com for assistance.';

  cardEl.appendChild(tooltip);

  setTimeout(() => {
    tooltip.style.opacity = '0';
    tooltip.style.transition = 'opacity 0.3s ease';
    setTimeout(() => tooltip.remove(), 300);
  }, 4000);
}

// Opportunity Card Three-Dot Actions
function handleOppAction(action, oppId) {
  const opp = opportunitiesData.find(o => o.id === oppId);
  if (!opp) return;

  if (action === 'view') {
    openOppDetailsModal(oppId);
  } else if (action === 'edit') {
    openEditOppModal(oppId);
  } else if (action === 'duplicate') {
    // Duplicate Opportunity in the same stage (Requirement 5)
    const newId = `OPP-${Math.floor(100 + Math.random() * 900)}`;
    const clonedOpp = JSON.parse(JSON.stringify(opp));
    clonedOpp.id = newId;
    clonedOpp.name = `${opp.name} (Copy)`;
    clonedOpp.createdDate = new Date().toISOString();
    opportunitiesData.unshift(clonedOpp);
    renderCurrentOpportunityView();
    showToast('Duplicated', `Created duplicate: "${clonedOpp.name}"`, 'success');
  } else if (action === 'delete') {
    // Delete Opportunity
    if (confirm(`Are you sure you want to delete opportunity "${opp.name}"?`)) {
      opportunitiesData = opportunitiesData.filter(o => o.id !== oppId);
      renderCurrentOpportunityView();
      showToast('Deleted', `Opportunity "${opp.name}" deleted.`, 'info');
    }
  } else if (action === 'inquiry') {
    showToast('Create Inquiry', `Inquiry generated for ${opp.name}`, 'success');
  } else if (action === 'quotation') {
    showToast('Create Quotation', `Quotation draft created for ${opp.name}`, 'success');
  }
}

// Opportunity Details Modal
function openOppDetailsModal(oppId) {
  const opp = opportunitiesData.find(o => o.id === oppId);
  if (!opp) return;

  const stage = opportunityStages.find(s => s.id === opp.stage);

  document.getElementById('detailOppName').textContent = opp.fullName || opp.name;
  document.getElementById('detailOppStage').innerHTML = `
    <span class="opp-stage-badge" style="background: rgba(${stage ? (stage.id === 'won' ? '22, 163, 74' : (stage.id === 'lost' ? '220, 38, 38' : '37, 99, 235')) : '100,100,100'}, 0.12); color: ${stage ? stage.color : '#333'};">
      ● ${stage ? stage.name : opp.stage}
    </span>
  `;
  document.getElementById('detailPartyName').textContent = opp.fullNameParty || opp.party;
  document.getElementById('detailTradeLane').textContent = opp.tradeLane;
  document.getElementById('detailExpRevenue').textContent = opp.expRevDisplay;
  document.getElementById('detailExpProfit').textContent = opp.expProfitDisplay;
  document.getElementById('detailClosureDate').textContent = opp.closureDate;
  document.getElementById('detailOppId').textContent = opp.id;

  // Render Audit History
  const auditContainer = document.getElementById('oppAuditLogList');
  if (auditContainer) {
    auditContainer.innerHTML = '';
    const history = opp.auditHistory || [];
    if (history.length === 0) {
      auditContainer.innerHTML = '<div style="color: #94a3b8; font-size: 11.5px;">No audit history recorded.</div>';
    } else {
      history.forEach(item => {
        const entry = document.createElement('div');
        entry.className = 'audit-log-entry';
        entry.innerHTML = `
          <span>${escapeHtml(item.text)}</span>
          <span class="audit-log-time">${escapeHtml(item.timestamp)}</span>
        `;
        auditContainer.appendChild(entry);
      });
    }
  }

  const modal = document.getElementById('oppDetailsModal');
  const backdrop = document.getElementById('oppDetailsBackdrop');
  if (modal) modal.classList.add('open');
  if (backdrop) backdrop.style.display = 'block';
}

function closeOppDetailsModal() {
  const modal = document.getElementById('oppDetailsModal');
  const backdrop = document.getElementById('oppDetailsBackdrop');
  if (modal) modal.classList.remove('open');
  if (backdrop) backdrop.style.display = 'none';
}

// Create & Edit Opportunity Modal Handlers
let editingOppId = null;

function openEditOppModal(oppId) {
  const opp = opportunitiesData.find(o => o.id === oppId);
  if (!opp) return;

  editingOppId = oppId;
  const modal = document.getElementById('createOppModal');
  const title = modal ? modal.querySelector('.opp-modal-title') : null;
  if (title) title.textContent = 'Edit Opportunity';

  document.getElementById('inputNewOppName').value = opp.name;
  document.getElementById('inputNewPartyName').value = opp.party;
  document.getElementById('inputNewTradeLane').value = opp.tradeLane;
  document.getElementById('inputNewExpRev').value = opp.expRevValue || '';
  document.getElementById('inputNewExpProfit').value = opp.expProfitValue || '';
  document.getElementById('selectNewStage').value = opp.stage;
  document.getElementById('inputNewClosure').value = '2026-08-31';

  if (modal) modal.classList.add('open');
  const backdrop = document.getElementById('oppDetailsBackdrop');
  if (backdrop) backdrop.style.display = 'block';
}

function closeCreateOppModal() {
  const modal = document.getElementById('createOppModal');
  const backdrop = document.getElementById('oppDetailsBackdrop');
  if (modal) modal.classList.remove('open');
  if (backdrop) backdrop.style.display = 'none';
  editingOppId = null;
  const title = modal ? modal.querySelector('.opp-modal-title') : null;
  if (title) title.textContent = 'Create Opportunity';
}

function handleCreateOppSubmit() {
  const nameInput = document.getElementById('inputNewOppName');
  const partyInput = document.getElementById('inputNewPartyName');
  const tradeInput = document.getElementById('inputNewTradeLane');
  const revInput = document.getElementById('inputNewExpRev');
  const profitInput = document.getElementById('inputNewExpProfit');
  const stageSelect = document.getElementById('selectNewStage');
  const closureInput = document.getElementById('inputNewClosure');

  const name = (nameInput.value || '').trim();
  const party = (partyInput.value || '').trim();
  if (!name || !party) {
    showToast('Validation Error', 'Please enter Opportunity Name and Party Name', 'warning');
    return;
  }

  const revVal = parseFloat(revInput.value) || 1500000;
  const profitVal = parseFloat(profitInput.value) || 350000;
  const stage = stageSelect.value || 'new';

  if (editingOppId) {
    const opp = opportunitiesData.find(o => o.id === editingOppId);
    if (opp) {
      opp.name = name;
      opp.fullName = name;
      opp.party = party;
      opp.tradeLane = tradeInput.value.trim() || opp.tradeLane;
      opp.expRevDisplay = formatIndianCurrency(revVal);
      opp.expRevValue = revVal;
      opp.expProfitDisplay = formatIndianCurrency(profitVal);
      opp.expProfitValue = profitVal;
      opp.closureDate = closureInput.value || opp.closureDate;
      opp.stage = stage;
      if (!opp.auditHistory) opp.auditHistory = [];
      opp.auditHistory.unshift({
        timestamp: new Date().toLocaleString(),
        text: `Opportunity details edited by Sakshi Barnwal`
      });
      showToast('Updated', `Opportunity "${opp.name}" updated successfully.`, 'success');
    }
  } else {
    const newOpp = {
      id: `OPP-${Math.floor(100 + Math.random() * 900)}`,
      name: name,
      fullName: name,
      party: party,
      tradeLane: tradeInput.value.trim() || 'India - Germany',
      expRevDisplay: formatIndianCurrency(revVal),
      expRevValue: revVal,
      expProfitDisplay: formatIndianCurrency(profitVal),
      expProfitValue: profitVal,
      closureDate: closureInput.value || '31-Aug-2026',
      stage: stage,
      createdDate: new Date().toISOString(),
      auditHistory: [
        { timestamp: new Date().toLocaleString(), text: `Opportunity created in ${stage} stage by Sakshi Barnwal` }
      ]
    };

    opportunitiesData.unshift(newOpp);
    showToast('Created', `Opportunity "${newOpp.name}" created successfully.`, 'success');
  }

  closeCreateOppModal();

  nameInput.value = '';
  partyInput.value = '';
  tradeInput.value = '';
  revInput.value = '';
  profitInput.value = '';

  renderCurrentOpportunityView();
}

// Render Opportunities Table View (Requirement 1 & 2)
function renderOppTableView() {
  const tbody = document.getElementById('oppTableBody');
  if (!tbody) return;

  tbody.innerHTML = '';

  let list = [...opportunitiesData];
  if (currentSearchQuery) {
    list = list.filter(o =>
      o.name.toLowerCase().includes(currentSearchQuery) ||
      o.party.toLowerCase().includes(currentSearchQuery) ||
      o.tradeLane.toLowerCase().includes(currentSearchQuery)
    );
  }

  list.forEach(opp => {
    const stage = opportunityStages.find(s => s.id === opp.stage);
    const tr = document.createElement('tr');

    tr.innerHTML = `
      <td style="width: 32px; text-align: center; color: #94a3b8;">⋮⋮</td>
      <td class="icon-cell" style="cursor: pointer;" title="Actions">⋮</td>
      <td class="icon-cell" style="color: #d1d5db; cursor: pointer;">☆</td>
      <td>
        <a class="lead-id-link opp-name-click" data-id="${opp.id}">${escapeHtml(opp.name)}</a>
      </td>
      <td>
        <div class="lead-company-name">${escapeHtml(opp.party)}</div>
      </td>
      <td>
        <span class="opp-stage-badge" style="background: rgba(${stage ? (stage.id === 'won' ? '22, 163, 74' : (stage.id === 'lost' ? '220, 38, 38' : '37, 99, 235')) : '100,100,100'}, 0.12); color: ${stage ? stage.color : '#333'};">
          ● ${stage ? stage.name : opp.stage}
        </span>
      </td>
      <td>${escapeHtml(opp.tradeLane)}</td>
      <td style="font-weight: 600; color: #1e293b;">${escapeHtml(opp.expRevDisplay)}</td>
      <td style="font-weight: 600; color: #12b76a;">${escapeHtml(opp.expProfitDisplay)}</td>
      <td>${escapeHtml(opp.closureDate || '--')}</td>
    `;

    tr.querySelector('.opp-name-click').addEventListener('click', () => {
      openOppDetailsModal(opp.id);
    });

    tbody.appendChild(tr);
  });
}

// Update Scroll Navigation Arrows (Requirement 3: Enable only when scroll is required)
function updateScrollArrows() {
  const container = document.getElementById('pipelineScrollContainer');
  const arrowLeft = document.getElementById('scrollArrowLeft');
  const arrowRight = document.getElementById('scrollArrowRight');

  if (!container || !arrowLeft || !arrowRight) return;

  const maxScroll = container.scrollWidth - container.clientWidth;
  const isOverflowing = maxScroll > 5;

  if (!isOverflowing) {
    arrowLeft.disabled = true;
    arrowRight.disabled = true;
    return;
  }

  arrowLeft.disabled = container.scrollLeft <= 5;
  arrowRight.disabled = container.scrollLeft >= maxScroll - 5;
}

// Toggle Opportunity Update Permission in Admin Modal (Interactive Screen 1)
function toggleOppUpdatePermission() {
  const btn = document.getElementById('btnPermOppUpdate');
  if (!btn) return;
  oppUpdatePermissionDraft = !oppUpdatePermissionDraft;
  btn.classList.toggle('checked', oppUpdatePermissionDraft);
}

// Commit Permissions from Modal
function saveOppPermissions() {
  oppUpdatePermissionEnabled = oppUpdatePermissionDraft;
  updateOppPermissionPillUI();
}

function updateOppPermissionPillUI() {
  const btn = document.getElementById('btnPermOppUpdate');
  if (btn) {
    btn.classList.toggle('checked', oppUpdatePermissionEnabled);
  }
}

// Global Prototype Testing Helpers
window.toggleActiveStages = function() {
  const allActive = opportunityStages.some(s => s.active);
  opportunityStages.forEach(s => s.active = !allActive);
  renderPipelineView();
  showToast(
    !allActive ? 'Stages Restored' : 'All Stages Deactivated',
    !allActive ? 'Active stages restored to default funnel.' : 'Simulating: No Active Opportunity Stages configured.',
    !allActive ? 'success' : 'warning'
  );
};

window.resetOpportunitiesPrototype = function() {
  opportunitiesData = JSON.parse(JSON.stringify(INITIAL_OPPORTUNITIES));
  oppUpdatePermissionEnabled = true;
  oppUpdatePermissionDraft = true;
  opportunityStages.forEach(s => s.active = true);
  updateOppPermissionPillUI();
  renderCurrentOpportunityView();
  showToast('Reset Complete', 'Opportunities data and permissions reset to original screenshot state.', 'info');
};

/* ========================================================================= */
/* CONTACTS MODULE & PARTY LOOKUP-WITH-CREATE (FOS Ticket Implementation)    */
/* ========================================================================= */

let partyDropdownClicked = false;

function initContactsModule() {
  // Populate Country dropdown in Address Capture Modal
  populateCountryDropdown();

  // Bind Create Contact Button in Header (Image 1 Match)
  const btnCreateContact = document.getElementById('btnCreateContact');
  if (btnCreateContact) {
    btnCreateContact.addEventListener('click', openContactDrawer);
  }

  // Bind Contact Drawer Controls (Image 2 Match)
  const btnCloseContactDrawer = document.getElementById('btnCloseContactDrawer');
  const btnCancelContact = document.getElementById('btnCancelContact');
  const contactDrawerBackdrop = document.getElementById('contactDrawerBackdrop');
  const btnToggleAccordion = document.getElementById('btnToggleAccordion');
  const btnDismissPermBanner = document.getElementById('btnDismissPermBanner');
  const btnDismissDupBanner = document.getElementById('btnDismissDupBanner');
  const btnSaveContact = document.getElementById('btnSaveContact');

  if (btnCloseContactDrawer) btnCloseContactDrawer.addEventListener('click', closeContactDrawer);
  if (btnCancelContact) btnCancelContact.addEventListener('click', closeContactDrawer);
  if (contactDrawerBackdrop) contactDrawerBackdrop.addEventListener('click', closeContactDrawer);
  if (btnDismissPermBanner) {
    btnDismissPermBanner.addEventListener('click', () => {
      const banner = document.getElementById('contactPermissionBanner');
      if (banner) banner.style.display = 'none';
    });
  }
  if (btnDismissDupBanner) {
    btnDismissDupBanner.addEventListener('click', () => {
      const banner = document.getElementById('contactDuplicateBanner');
      if (banner) banner.style.display = 'none';
    });
  }
  if (btnSaveContact) btnSaveContact.addEventListener('click', handleSaveContact);

  // Bind Party Lookup Input Controls
  const contactPartyInput = document.getElementById('contactPartyInput');
  const btnPartyClear = document.getElementById('btnPartyClear');
  const btnEditCapturedAddress = document.getElementById('btnEditCapturedAddress');

  if (contactPartyInput) {
    contactPartyInput.addEventListener('input', handlePartyInput);
    contactPartyInput.addEventListener('focus', handlePartyFocus);
    contactPartyInput.addEventListener('blur', handlePartyBlur);
    contactPartyInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const partyDropdownPanel = document.getElementById('partyDropdownPanel');
        if (partyDropdownPanel && partyDropdownPanel.style.display !== 'none') {
          const addOption = document.getElementById('partyAddLeadOption');
          if (addOption) {
            addOption.click();
          }
        }
      }
    });
  }

  if (btnPartyClear) {
    btnPartyClear.addEventListener('click', handleClearPartyInput);
  }

  if (btnEditCapturedAddress) {
    btnEditCapturedAddress.addEventListener('click', () => {
      const partyName = pendingNewParty ? pendingNewParty.companyName : (currentTypedPartyValue.trim() || 'New Lead');
      openAddressCaptureModal(partyName);
    });
  }

  // Bind Address Capture Modal Controls
  const btnCloseAddressCapture = document.getElementById('btnCloseAddressCapture');
  const btnCancelAddressCapture = document.getElementById('btnCancelAddressCapture');
  const btnSaveAddressCapture = document.getElementById('btnSaveAddressCapture');
  const addressCaptureBackdrop = document.getElementById('addressCaptureBackdrop');
  const addrCaptureAddress = document.getElementById('addrCaptureAddress');

  if (btnCloseAddressCapture) btnCloseAddressCapture.addEventListener('click', handleCancelAddressCapture);
  if (btnCancelAddressCapture) btnCancelAddressCapture.addEventListener('click', handleCancelAddressCapture);
  if (addressCaptureBackdrop) addressCaptureBackdrop.addEventListener('click', handleCancelAddressCapture);
  if (btnSaveAddressCapture) btnSaveAddressCapture.addEventListener('click', handleSaveAddressCapture);

  if (addrCaptureAddress) {
    addrCaptureAddress.addEventListener('input', (e) => {
      const counter = document.getElementById('charCountAddress');
      if (counter) counter.textContent = `${e.target.value.length} / 500`;
    });
  }

  // Bind AC Test Controls Bar
  setupAcDemoControls();

  // Close Party dropdown on outside click
  document.addEventListener('click', (e) => {
    const partyGroupWrapper = document.getElementById('partyGroupWrapper');
    if (partyGroupWrapper && !partyGroupWrapper.contains(e.target)) {
      const partyDropdownPanel = document.getElementById('partyDropdownPanel');
      if (partyDropdownPanel) partyDropdownPanel.style.display = 'none';
    }
  });

  // Global Keyboard Navigation: Escape key handler (AC-2.2 & AC-4.3)
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      const addrModal = document.getElementById('addressCaptureModal');
      if (addrModal && addrModal.style.display !== 'none') {
        handleCancelAddressCapture();
        return;
      }
      const contactDrawer = document.getElementById('contactDrawer');
      if (contactDrawer && contactDrawer.classList.contains('open')) {
        closeContactDrawer();
      }
    }
  });

  // Initial table render
  renderContactsTable();
  updateLeadsCreatePermissionUI();
}

// Populate Country Master Options into Select
function populateCountryDropdown() {
  const select = document.getElementById('addrCaptureCountry');
  if (!select) return;

  // Preserve first option
  select.innerHTML = '<option value="">--Select Country--</option>';

  // Sort country entries alphabetically by name
  const sortedEntries = Object.entries(COUNTRY_MASTER).sort((a, b) => a[1].localeCompare(b[1]));

  sortedEntries.forEach(([code, name]) => {
    const opt = document.createElement('option');
    opt.value = code;
    opt.textContent = `${name} (${code})`;
    select.appendChild(opt);
  });
}

// Render Contacts Table View (Exact Replica of Image 1)
function renderContactsTable() {
  const tbody = document.getElementById('contactsTableBody');
  const countEl = document.getElementById('contactsResultsCount');
  if (!tbody) return;

  tbody.innerHTML = '';

  let list = [...contactsData];

  // Search Filter
  if (contactSearchQuery) {
    const q = contactSearchQuery.toLowerCase();
    list = list.filter(c =>
      c.name.toLowerCase().includes(q) ||
      c.partyName.toLowerCase().includes(q) ||
      (c.partyLocation && c.partyLocation.toLowerCase().includes(q)) ||
      c.owner.toLowerCase().includes(q) ||
      c.email.toLowerCase().includes(q) ||
      (c.phone && c.phone.toLowerCase().includes(q)) ||
      c.createdBy.toLowerCase().includes(q)
    );
  }

  if (countEl) {
    countEl.textContent = `of ${list.length} results`;
  }

  if (list.length === 0) {
    const emptyRow = document.createElement('tr');
    emptyRow.innerHTML = `
      <td colspan="9" style="text-align: center; padding: 36px 12px; color: #94a3b8;">
        <div style="font-size: 14px; font-weight: 600; margin-bottom: 4px;">No contacts found</div>
        <div style="font-size: 12px;">Try adjusting your search criteria or click "+ Create Contact"</div>
      </td>
    `;
    tbody.appendChild(emptyRow);
    return;
  }

  list.forEach(contact => {
    const tr = document.createElement('tr');

    // Is there an "Updated" tag on party name? (Image 1 Match: "ABC Freight LLP Updated")
    const isUpdatedParty = contact.partyName.includes('Updated');
    const displayPartyName = isUpdatedParty ? contact.partyName.replace('Updated', '').trim() : contact.partyName;

    tr.innerHTML = `
      <td style="width: 32px; text-align: center; color: #94a3b8;">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="#9ca3af">
          <path d="M9 5a2 2 0 1 1-4 0 2 2 0 0 1 4 0zm0 7a2 2 0 1 1-4 0 2 2 0 0 1 4 0zm0 7a2 2 0 1 1-4 0 2 2 0 0 1 4 0zm10-14a2 2 0 1 1-4 0 2 2 0 0 1 4 0zm0 7a2 2 0 1 1-4 0 2 2 0 0 1 4 0zm0 7a2 2 0 1 1-4 0 2 2 0 0 1 4 0z" />
        </svg>
      </td>
      <td>
        <div class="contact-name-cell" title="${escapeHtml(contact.name)}">${escapeHtml(contact.name)}</div>
      </td>
      <td>
        <div class="contact-party-cell">
          <div class="contact-party-name">
            ${escapeHtml(displayPartyName)}
            ${isUpdatedParty ? '<span class="contact-party-tag">Updated</span>' : ''}
          </div>
          <div class="contact-party-location">${escapeHtml(contact.partyLocation || '--')}</div>
        </div>
      </td>
      <td>${escapeHtml(contact.owner || '--')}</td>
      <td>
        <a href="mailto:${escapeHtml(contact.email)}" style="color: #2563eb; text-decoration: none;">
          ${escapeHtml(contact.email || '--')}
        </a>
      </td>
      <td>${escapeHtml(contact.phone || '--')}</td>
      <td>${escapeHtml(contact.createdBy || '--')}</td>
      <td>${escapeHtml(contact.createdDate || '--')}</td>
    `;

    tbody.appendChild(tr);
  });
}

// Filter Leads Table Helper
function filterLeadsTable(query) {
  const tbody = document.getElementById('leadsTableBody');
  if (!tbody) return;

  tbody.innerHTML = '';
  let list = [...leadsData];
  if (query) {
    const q = query.toLowerCase();
    list = list.filter(l =>
      l.companyName.toLowerCase().includes(q) ||
      l.leadOwner.toLowerCase().includes(q) ||
      l.contactName.toLowerCase().includes(q)
    );
  }

  list.forEach(lead => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="width: 32px; text-align: center; color: #94a3b8;">⋮⋮</td>
      <td class="icon-cell" style="cursor: pointer;">⋮</td>
      <td class="icon-cell" style="color: #d1d5db;">☆</td>
      <td><a class="lead-id-link">${escapeHtml(lead.leadId)}</a></td>
      <td><div class="lead-company-name">${escapeHtml(lead.companyName)}</div></td>
      <td><span class="stage-badge stage-${(lead.stage || 'new').toLowerCase().replace(/\s+/g, '-')}">● ${escapeHtml(lead.stage)}</span></td>
      <td>${escapeHtml(lead.leadOwner)}</td>
      <td style="text-align: center;">${lead.temperature || '❄️'}</td>
      <td>${escapeHtml(lead.contactName)}</td>
      <td>${escapeHtml(lead.contactPhone)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Open Contact Drawer (Image 2 Match)
function openContactDrawer() {
  const drawer = document.getElementById('contactDrawer');
  const backdrop = document.getElementById('contactDrawerBackdrop');
  if (!drawer || !backdrop) return;

  resetContactForm();

  drawer.classList.add('open');
  backdrop.classList.add('open');

  // Auto-focus First Name input
  setTimeout(() => {
    const firstInput = document.getElementById('contactFirstName');
    if (firstInput) firstInput.focus();
  }, 200);
}

// Close Contact Drawer (AC-4.3: pending new party is discarded without persistence)
function closeContactDrawer() {
  const drawer = document.getElementById('contactDrawer');
  const backdrop = document.getElementById('contactDrawerBackdrop');
  if (drawer) drawer.classList.remove('open');
  if (backdrop) backdrop.classList.remove('open');

  // Clear uncommitted pending Lead data (AC-4.3)
  pendingNewParty = null;
  selectedExistingParty = null;
  lastAddressCapturePartyName = '';
  currentTypedPartyValue = '';
}

// Reset Contact Drawer Form Fields
function resetContactForm() {
  const fields = [
    'contactFirstName',
    'contactLastName',
    'contactEmail',
    'contactPhone',
    'contactPartyInput',
    'contactNotes'
  ];

  fields.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  const ownerSelect = document.getElementById('contactOwner');
  if (ownerSelect) ownerSelect.value = 'Sakshi Barnwal';

  const sourceSelect = document.getElementById('contactSource');
  if (sourceSelect) sourceSelect.value = 'Advertisement';

  // Hide errors and banners
  hideAllContactFieldErrors();

  const permBanner = document.getElementById('contactPermissionBanner');
  if (permBanner) permBanner.style.display = 'none';

  const dupBanner = document.getElementById('contactDuplicateBanner');
  if (dupBanner) dupBanner.style.display = 'none';

  const inlineInfo = document.getElementById('partyInlineInfoBox');
  if (inlineInfo) inlineInfo.style.display = 'none';

  const pendingBadge = document.getElementById('partyPendingBadge');
  if (pendingBadge) pendingBadge.style.display = 'none';

  const clearBtn = document.getElementById('btnPartyClear');
  if (clearBtn) clearBtn.style.display = 'none';

  const partyWarn = document.getElementById('partySearchWarning');
  if (partyWarn) partyWarn.style.display = 'none';

  const partyDropdownPanel = document.getElementById('partyDropdownPanel');
  if (partyDropdownPanel) partyDropdownPanel.style.display = 'none';

  pendingNewParty = null;
  selectedExistingParty = null;
  currentTypedPartyValue = '';
  lastAddressCapturePartyName = '';
}

function hideAllContactFieldErrors() {
  document.querySelectorAll('.contact-drawer .fos-field-error').forEach(e => {
    e.classList.remove('visible');
    e.textContent = '';
  });
  document.querySelectorAll('.contact-drawer .fos-form-input, .contact-drawer .fos-form-select').forEach(i => {
    i.classList.remove('has-error');
  });
}


// Clear Party Input handler
function handleClearPartyInput() {
  const input = document.getElementById('contactPartyInput');
  const clearBtn = document.getElementById('btnPartyClear');
  const pendingBadge = document.getElementById('partyPendingBadge');
  const inlineInfo = document.getElementById('partyInlineInfoBox');
  const dropdown = document.getElementById('partyDropdownPanel');

  if (input) input.value = '';
  if (clearBtn) clearBtn.style.display = 'none';
  if (pendingBadge) pendingBadge.style.display = 'none';
  if (inlineInfo) inlineInfo.style.display = 'none';
  if (dropdown) dropdown.style.display = 'none';

  // AC-5.4: Clearing the field completely discards captured address
  pendingNewParty = null;
  selectedExistingParty = null;
  lastAddressCapturePartyName = '';
  currentTypedPartyValue = '';

  if (input) input.focus();
}

// Party Input Event Handler (Search as you type & detecting unmatched values)
function handlePartyInput(e) {
  const val = e.target.value;
  currentTypedPartyValue = val;

  const clearBtn = document.getElementById('btnPartyClear');
  if (clearBtn) {
    clearBtn.style.display = val.length > 0 ? 'flex' : 'none';
  }

  // If user clears completely (AC-5.4)
  if (!val.trim()) {
    pendingNewParty = null;
    selectedExistingParty = null;
    lastAddressCapturePartyName = '';
    const pendingBadge = document.getElementById('partyPendingBadge');
    if (pendingBadge) pendingBadge.style.display = 'none';
    const inlineInfo = document.getElementById('partyInlineInfoBox');
    if (inlineInfo) inlineInfo.style.display = 'none';
    const dropdown = document.getElementById('partyDropdownPanel');
    if (dropdown) dropdown.style.display = 'none';
    return;
  }

  renderPartyDropdown(val);
}

// Party Focus Event Handler
function handlePartyFocus(e) {
  if (e.target.value.trim()) {
    renderPartyDropdown(e.target.value);
  }
}

// Party Blur / Defocus Event Handler (AC-2.1, AC-2.4, AC-2.5, AC-2.6, AC-5.1)
function handlePartyBlur() {
  // Give dropdown item click handler priority
  setTimeout(() => {
    if (partyDropdownClicked) {
      partyDropdownClicked = false;
      return;
    }

    const input = document.getElementById('contactPartyInput');
    const dropdown = document.getElementById('partyDropdownPanel');
    if (dropdown) dropdown.style.display = 'none';

    if (!input) return;
    const typed = input.value.trim();
    if (!typed) return;

    // AC-5.1: Case/whitespace-trimmed match check against Active Parties
    const exactMatch = ACTIVE_PARTIES.find(p => p.name.trim().toLowerCase() === typed.toLowerCase());
    if (exactMatch) {
      // Matched existing party: Link it, do NOT create new Lead, do NOT open popup (AC-2.4 & AC-5.1)
      selectedExistingParty = exactMatch;
      pendingNewParty = null;
      lastAddressCapturePartyName = '';
      input.value = exactMatch.name;

      const pendingBadge = document.getElementById('partyPendingBadge');
      if (pendingBadge) pendingBadge.style.display = 'none';
      const inlineInfo = document.getElementById('partyInlineInfoBox');
      if (inlineInfo) inlineInfo.style.display = 'none';
      return;
    }

    // AC-2.5: If user already completed Address popup for this party and just corrected a typo
    if (pendingNewParty && lastAddressCapturePartyName &&
        (typed.toLowerCase().includes(lastAddressCapturePartyName.toLowerCase()) ||
         lastAddressCapturePartyName.toLowerCase().includes(typed.toLowerCase()))) {
      pendingNewParty.companyName = typed;
      const pendingBadge = document.getElementById('partyPendingBadge');
      if (pendingBadge) pendingBadge.style.display = 'block';
      const inlineInfo = document.getElementById('partyInlineInfoBox');
      const inlineText = document.getElementById('partyInlineInfoText');
      const editBtn = document.getElementById('btnEditCapturedAddress');
      if (inlineInfo && inlineText) {
        inlineText.textContent = `"${typed}" has been added as pending lead`;
        inlineInfo.style.display = 'flex';
      }
      if (editBtn) {
        editBtn.textContent = 'Edit location';
        editBtn.style.display = 'inline-block';
      }
      return;
    }

    // Unmatched Value Defocus -> Open Address Capture popup directly (inline notification appears after entering address)
    const inlineInfo = document.getElementById('partyInlineInfoBox');
    if (inlineInfo) inlineInfo.style.display = 'none';

    // Open Address Capture Popup immediately!
    openAddressCaptureModal(typed);
  }, 220);
}

// Render Autocomplete Dropdown List
function renderPartyDropdown(query) {
  const panel = document.getElementById('partyDropdownPanel');
  const listContainer = document.getElementById('partyDropdownList');
  const warningEl = document.getElementById('partySearchWarning');
  if (!panel || !listContainer) return;

  const q = query.trim().toLowerCase();

  // AC-5.5: Simulated Search Timeout / Error
  if (simulatePartySearchTimeout) {
    if (warningEl) warningEl.style.display = 'block';
    listContainer.innerHTML = `
      <div class="party-add-lead-item" id="partyAddLeadOption">
        <span class="party-add-icon-badge">+</span>
        <span>Add "<strong>${escapeHtml(query.trim())}</strong>" as a new Lead</span>
      </div>
    `;
    const addBtn = document.getElementById('partyAddLeadOption');
    if (addBtn) {
      addBtn.addEventListener('mousedown', () => {
        partyDropdownClicked = true;
        panel.style.display = 'none';
        openAddressCaptureModal(query.trim());
      });
    }
    panel.style.display = 'block';
    return;
  }

  if (warningEl) warningEl.style.display = 'none';

  // Filter Active Parties (Vendor & Inactive are excluded per AC-5.2 & AC-5.3)
  const matches = ACTIVE_PARTIES.filter(p =>
    p.name.toLowerCase().includes(q) ||
    p.city.toLowerCase().includes(q) ||
    p.country.toLowerCase().includes(q)
  );

  const exactMatch = ACTIVE_PARTIES.find(p => p.name.trim().toLowerCase() === q);

  listContainer.innerHTML = '';

  if (matches.length > 0) {
    matches.forEach(party => {
      const item = document.createElement('div');
      item.className = 'party-dropdown-item';
      item.innerHTML = `
        <span class="party-item-title">${escapeHtml(party.name)} (${escapeHtml(party.city)}, ${escapeHtml(party.country)})</span>
        <span class="party-item-role-tag">${escapeHtml(party.role)}</span>
      `;
      item.addEventListener('mousedown', () => {
        partyDropdownClicked = true;
        selectExistingPartyOption(party);
      });
      listContainer.appendChild(item);
    });
  }

  // If NOT an exact match, display "+ Add '{typed}' as a new Lead" (Requirement 2 & Section 2)
  if (!exactMatch && query.trim().length > 0) {
    const addItem = document.createElement('div');
    addItem.className = 'party-add-lead-item';
    addItem.id = 'partyAddLeadOption';
    addItem.innerHTML = `
      <span class="party-add-icon-badge">+</span>
      <span>Add "<strong>${escapeHtml(query.trim())}</strong>" as a new Lead</span>
    `;
    addItem.addEventListener('mousedown', () => {
      partyDropdownClicked = true;
      panel.style.display = 'none';
      const inlineInfo = document.getElementById('partyInlineInfoBox');
      if (inlineInfo) inlineInfo.style.display = 'none';
      openAddressCaptureModal(query.trim());
    });
    listContainer.appendChild(addItem);
  }

  panel.style.display = (matches.length > 0 || !exactMatch) ? 'block' : 'none';
}

// Select Existing Party Option (AC-2.4)
function selectExistingPartyOption(party) {
  const input = document.getElementById('contactPartyInput');
  const panel = document.getElementById('partyDropdownPanel');
  const inlineInfo = document.getElementById('partyInlineInfoBox');
  const pendingBadge = document.getElementById('partyPendingBadge');
  const clearBtn = document.getElementById('btnPartyClear');
  const partyErr = document.getElementById('errContactParty');

  if (input) input.value = party.name;
  if (panel) panel.style.display = 'none';
  if (inlineInfo) inlineInfo.style.display = 'none';
  if (pendingBadge) pendingBadge.style.display = 'none';
  if (clearBtn) clearBtn.style.display = 'flex';
  if (partyErr) partyErr.classList.remove('visible');

  selectedExistingParty = party;
  pendingNewParty = null;
  lastAddressCapturePartyName = '';
  currentTypedPartyValue = party.name;
}

// Open Address Capture Modal (Section 3: Modal over Contact Drawer)
function openAddressCaptureModal(partyName) {
  const modal = document.getElementById('addressCaptureModal');
  const backdrop = document.getElementById('addressCaptureBackdrop');
  const title = document.getElementById('addressCaptureTitle');
  const subtitle = document.getElementById('addressModalSubtitle');
  if (!modal || !backdrop) return;

  if (title) title.textContent = `Address Details`;
  if (subtitle) {
    subtitle.textContent = `Please provide location details to create "${partyName}" as a new Lead.`;
  }

  // Clear modal error messages
  clearAddressModalErrors();

  const addrInput = document.getElementById('addrCaptureAddress');
  const cityInput = document.getElementById('addrCaptureCity');
  const countrySelect = document.getElementById('addrCaptureCountry');
  const counter = document.getElementById('charCountAddress');

  if (pendingNewParty) {
    if (addrInput) addrInput.value = pendingNewParty.address || '';
    if (cityInput) cityInput.value = pendingNewParty.city || '';
    if (countrySelect) countrySelect.value = pendingNewParty.countryCode || '';
    if (counter) counter.textContent = `${(pendingNewParty.address || '').length} / 500`;
  } else {
    if (addrInput) addrInput.value = '';
    if (cityInput) cityInput.value = '';
    if (countrySelect) countrySelect.value = '';
    if (counter) counter.textContent = '0 / 500';
  }

  modal.style.display = 'flex';
  backdrop.style.display = 'block';

  // Auto-focus Address Line 1
  setTimeout(() => {
    if (addrInput) addrInput.focus();
  }, 150);
}

// Cancel Address Capture Popup (AC-2.2: Clears Party field to empty!)
function handleCancelAddressCapture() {
  const modal = document.getElementById('addressCaptureModal');
  const backdrop = document.getElementById('addressCaptureBackdrop');
  if (modal) modal.style.display = 'none';
  if (backdrop) backdrop.style.display = 'none';

  // AC-2.2: The Party field is cleared back to empty — a half-finished pending state is NOT allowed
  if (!pendingNewParty) {
    const partyInput = document.getElementById('contactPartyInput');
    const inlineInfo = document.getElementById('partyInlineInfoBox');
    const pendingBadge = document.getElementById('partyPendingBadge');
    const clearBtn = document.getElementById('btnPartyClear');

    if (partyInput) partyInput.value = '';
    if (inlineInfo) inlineInfo.style.display = 'none';
    if (pendingBadge) pendingBadge.style.display = 'none';
    if (clearBtn) clearBtn.style.display = 'none';

    currentTypedPartyValue = '';
    lastAddressCapturePartyName = '';
  }
}

function clearAddressModalErrors() {
  ['errAddrCaptureAddress', 'errAddrCaptureCity', 'errAddrCaptureCountry'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.classList.remove('visible');
      el.textContent = '';
    }
  });

  ['addrCaptureAddress', 'addrCaptureCity', 'addrCaptureCountry'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('has-error');
  });
}

// Save Address Capture Popup (AC-2.3 & AC-3.1 to AC-3.5)
function handleSaveAddressCapture() {
  clearAddressModalErrors();

  const addrInput = document.getElementById('addrCaptureAddress');
  const cityInput = document.getElementById('addrCaptureCity');
  const countrySelect = document.getElementById('addrCaptureCountry');

  const address = addrInput ? addrInput.value.trim() : '';
  const city = cityInput ? cityInput.value.trim() : '';
  const countryCode = countrySelect ? countrySelect.value : '';
  const countryName = COUNTRY_MASTER[countryCode] || '';

  let hasError = false;

  // AC-3.1 & AC-3.4: Address validations
  const errAddr = document.getElementById('errAddrCaptureAddress');
  if (!address) {
    if (errAddr) {
      errAddr.textContent = 'Address is missing';
      errAddr.classList.add('visible');
    }
    if (addrInput) addrInput.classList.add('has-error');
    hasError = true;
  } else if (address.length > 500) {
    if (errAddr) {
      errAddr.textContent = 'Address cannot exceed 500 characters';
      errAddr.classList.add('visible');
    }
    if (addrInput) addrInput.classList.add('has-error');
    hasError = true;
  }

  // AC-3.2 & AC-3.5: City validations
  const errCity = document.getElementById('errAddrCaptureCity');
  if (!city) {
    if (errCity) {
      errCity.textContent = 'City is missing';
      errCity.classList.add('visible');
    }
    if (cityInput) cityInput.classList.add('has-error');
    hasError = true;
  } else if (city.length > 100) {
    if (errCity) {
      errCity.textContent = 'City cannot exceed 100 characters';
      errCity.classList.add('visible');
    }
    if (cityInput) cityInput.classList.add('has-error');
    hasError = true;
  }

  // AC-3.3: Country validation
  const errCountry = document.getElementById('errAddrCaptureCountry');
  if (!countryCode) {
    if (errCountry) {
      errCountry.textContent = 'Country is missing';
      errCountry.classList.add('visible');
    }
    if (countrySelect) countrySelect.classList.add('has-error');
    hasError = true;
  }

  if (hasError) return;

  // Success: Attach captured location to pending new Party
  const partyInput = document.getElementById('contactPartyInput');
  const companyName = currentTypedPartyValue.trim();

  pendingNewParty = {
    companyName: companyName,
    address: address,
    city: city,
    countryCode: countryCode,
    countryName: countryName
  };
  lastAddressCapturePartyName = companyName;

  // Close Popup
  const modal = document.getElementById('addressCaptureModal');
  const backdrop = document.getElementById('addressCaptureBackdrop');
  if (modal) modal.style.display = 'none';
  if (backdrop) backdrop.style.display = 'none';

  // Update Contact Drawer UI (AC-2.3)
  if (partyInput) partyInput.value = companyName;

  const pendingBadge = document.getElementById('partyPendingBadge');
  if (pendingBadge) pendingBadge.style.display = 'block';

  const inlineInfo = document.getElementById('partyInlineInfoBox');
  const inlineText = document.getElementById('partyInlineInfoText');
  const editBtn = document.getElementById('btnEditCapturedAddress');

  if (inlineInfo && inlineText) {
    inlineText.textContent = `"${companyName}" has been added as pending lead`;
    inlineInfo.style.display = 'flex';
  }
  if (editBtn) {
    editBtn.textContent = 'Edit location';
    editBtn.style.display = 'inline-block';
  }

  const partyErr = document.getElementById('errContactParty');
  if (partyErr) partyErr.classList.remove('visible');

  showToast(
    'Location Captured',
    `Address details attached to pending Lead "${companyName}".`,
    'info',
    3000
  );
}

// Contact Form Save Handler (AC-4.1, AC-4.2, AC-5.6)
function handleSaveContact() {
  hideAllContactFieldErrors();

  const permBanner = document.getElementById('contactPermissionBanner');
  const dupBanner = document.getElementById('contactDuplicateBanner');
  if (permBanner) permBanner.style.display = 'none';
  if (dupBanner) dupBanner.style.display = 'none';

  const firstNameInput = document.getElementById('contactFirstName');
  const lastNameInput = document.getElementById('contactLastName');
  const emailInput = document.getElementById('contactEmail');
  const phoneInput = document.getElementById('contactPhone');
  const ownerSelect = document.getElementById('contactOwner');
  const sourceSelect = document.getElementById('contactSource');
  const partyInput = document.getElementById('contactPartyInput');
  const notesInput = document.getElementById('contactNotes');

  const firstName = firstNameInput ? firstNameInput.value.trim() : '';
  const lastName = lastNameInput ? lastNameInput.value.trim() : '';
  const email = emailInput ? emailInput.value.trim() : '';
  const phone = phoneInput ? phoneInput.value.trim() : '';
  const owner = ownerSelect ? ownerSelect.value : 'Sakshi Barnwal';
  const source = sourceSelect ? sourceSelect.value : 'Advertisement';
  const partyVal = partyInput ? partyInput.value.trim() : '';
  const notes = notesInput ? notesInput.value.trim() : '';

  let hasError = false;

  // Validate First Name *
  if (!firstName) {
    const err = document.getElementById('errContactFirstName');
    if (err) {
      err.textContent = 'First Name is missing';
      err.classList.add('visible');
    }
    if (firstNameInput) firstNameInput.classList.add('has-error');
    hasError = true;
  }

  // Validate Last Name *
  if (!lastName) {
    const err = document.getElementById('errContactLastName');
    if (err) {
      err.textContent = 'Last Name is missing';
      err.classList.add('visible');
    }
    if (lastNameInput) lastNameInput.classList.add('has-error');
    hasError = true;
  }

  // Validate Email *
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!email) {
    const err = document.getElementById('errContactEmail');
    if (err) {
      err.textContent = 'Email is missing';
      err.classList.add('visible');
    }
    if (emailInput) emailInput.classList.add('has-error');
    hasError = true;
  } else if (!emailRegex.test(email)) {
    const err = document.getElementById('errContactEmail');
    if (err) {
      err.textContent = 'Invalid email address format';
      err.classList.add('visible');
    }
    if (emailInput) emailInput.classList.add('has-error');
    hasError = true;
  }

  // Validate Party *
  if (!partyVal) {
    const err = document.getElementById('errContactParty');
    if (err) {
      err.textContent = 'Party is missing';
      err.classList.add('visible');
    }
    if (partyInput) partyInput.classList.add('has-error');
    hasError = true;
  }

  if (hasError) return;

  // If party is typed but user hasn't completed Address Capture modal
  if (!selectedExistingParty && !pendingNewParty) {
    // Defocus / open popup for this value
    openAddressCaptureModal(partyVal);
    return;
  }

  // Check Company Name Length validation for new Party: 3-100 characters (Ticket Spec)
  if (pendingNewParty) {
    if (pendingNewParty.companyName.length < 3 || pendingNewParty.companyName.length > 100) {
      const err = document.getElementById('errContactParty');
      if (err) {
        err.textContent = 'New Party Name must be between 3 and 100 characters';
        err.classList.add('visible');
      }
      return;
    }

    // AC-5.6: Permission check — Leads -> Create permission required for implicit lead creation
    if (!leadsCreatePermissionEnabled) {
      if (permBanner) {
        permBanner.style.display = 'flex';
      }
      showToast(
        'Permission Denied',
        'You do not have permission to create a new party/lead. Please contact your Admin or service@freightoscope.com for assistance.',
        'error',
        6500
      );
      // Scroll drawer to top so user sees the permission error banner
      const drawerBody = document.querySelector('.contact-drawer-body');
      if (drawerBody) drawerBody.scrollTop = 0;
      return;
    }
  }

  // AC-4.2: Duplicate Email Check (Atomicity Failure Test)
  const isDuplicateEmail = contactsData.some(c => c.email.toLowerCase() === email.toLowerCase());
  if (isDuplicateEmail) {
    if (dupBanner) {
      dupBanner.style.display = 'flex';
    }
    const errEmail = document.getElementById('errContactEmail');
    if (errEmail) {
      errEmail.textContent = 'Email already exists for another contact';
      errEmail.classList.add('visible');
    }
    if (emailInput) emailInput.classList.add('has-error');

    showToast(
      'Duplicate Email',
      `A contact with email "${email}" already exists. Pending Lead was NOT created.`,
      'error',
      5500
    );

    // Scroll to top of drawer
    const drawerBody = document.querySelector('.contact-drawer-body');
    if (drawerBody) drawerBody.scrollTop = 0;

    // AC-4.2: Pending Lead is NOT persisted; Party field retains pending state for retry
    return;
  }

  // =========================================================================
  // AC-4.1: ATOMIC COMMIT ON SUCCESS (Both Lead and Contact created together)
  // =========================================================================
  let finalPartyName = partyVal;
  let finalPartyLocation = '';

  if (pendingNewParty) {
    finalPartyName = pendingNewParty.companyName;
    finalPartyLocation = `${pendingNewParty.city}, ${pendingNewParty.countryName}`;

    // AC-1.1: Lead Owner Email Id = currently logged-in user's email (sakshi.barnwal@freightoscope.com)
    // AC-1.2 & AC-5.7: Lead Stage = "New" (or fallback first stage if "New" stage master missing)
    // AC-1.3: Party Role = "Lead"
    const fallbackStage = LEAD_STAGES.includes('New') ? 'New' : LEAD_STAGES[0];

    const newLeadRecord = {
      leadId: `LEAD-${Math.floor(1000 + Math.random() * 9000)}`,
      companyName: pendingNewParty.companyName,
      location: finalPartyLocation,
      stage: fallbackStage,
      leadOwner: 'Sakshi Barnwal',
      leadOwnerEmail: 'sakshi.barnwal@freightoscope.com',
      temperature: '❄️',
      contactName: `${firstName} ${lastName}`,
      contactPhone: phone || '--',
      partyRole: 'Lead',
      addressLine1: pendingNewParty.address,
      city: pendingNewParty.city,
      countryCode: pendingNewParty.countryCode,
      createdDate: new Date().toISOString()
    };

    // Commit new Lead to leads database
    leadsData.unshift(newLeadRecord);

    // Also register in ACTIVE_PARTIES so subsequent searches immediately find it
    ACTIVE_PARTIES.unshift({
      name: pendingNewParty.companyName,
      city: pendingNewParty.city,
      country: pendingNewParty.countryName,
      countryCode: pendingNewParty.countryCode,
      role: 'Lead',
      location: finalPartyLocation
    });

    filterLeadsTable('');
  } else if (selectedExistingParty) {
    finalPartyName = selectedExistingParty.name;
    finalPartyLocation = `${selectedExistingParty.city}, ${selectedExistingParty.country}`;
  }

  // Format Current Date and Time (e.g., 9/16/26, 6:30 PM)
  const now = new Date();
  const dateStr = `${now.getMonth() + 1}/${now.getDate()}/${String(now.getFullYear()).slice(2)}, ${now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;

  const newContactRecord = {
    id: `CNT-${String(contactsData.length + 1).padStart(3, '0')}`,
    name: `${firstName} ${lastName}`,
    partyName: finalPartyName,
    partyLocation: finalPartyLocation,
    owner: owner,
    email: email,
    phone: phone || '--',
    createdBy: 'Sakshi Barnwal',
    createdDate: dateStr,
    department: '',
    designation: '',
    secondaryPhone: '',
    notes: notes
  };

  // Commit new Contact
  contactsData.unshift(newContactRecord);

  // Close Contact Drawer
  closeContactDrawer();

  // Re-render Contacts Table
  renderContactsTable();

  // Success Notification
  showToast(
    'Contact Created Successfully',
    pendingNewParty
      ? `Contact "${newContactRecord.name}" created and new Lead "${finalPartyName}" auto-created in stage "New"!`
      : `Contact "${newContactRecord.name}" created and linked to "${finalPartyName}".`,
    'success',
    5000
  );
}

// Toggle Leads Create Permission (AC-5.6)
function toggleLeadsCreatePermission(showNotice = false) {
  leadsCreatePermissionEnabled = !leadsCreatePermissionEnabled;
  updateLeadsCreatePermissionUI();

  if (showNotice) {
    showToast(
      leadsCreatePermissionEnabled ? 'Permission Granted' : 'Permission Revoked',
      leadsCreatePermissionEnabled
        ? 'Sales -> Leads -> Create permission is ENABLED.'
        : 'Sales -> Leads -> Create permission is DISABLED (Contact Save will block background Lead creation per AC-5.6).',
      leadsCreatePermissionEnabled ? 'success' : 'warning',
      4000
    );
  }
}

function updateLeadsCreatePermissionUI() {
  const btn = document.getElementById('btnPermLeadAdd');
  const box = document.getElementById('boxPermLeadAdd');
  const statusText = document.getElementById('permLeadAddStatusText');
  const toggleBtn = document.getElementById('btnTogglePermLeadAdd');

  if (btn) btn.classList.toggle('checked', leadsCreatePermissionEnabled);
  if (box) box.classList.toggle('checked', leadsCreatePermissionEnabled);

  if (statusText) {
    statusText.textContent = `Leads Create Perm: [${leadsCreatePermissionEnabled ? 'ENABLED' : 'DISABLED'}]`;
  }
  if (toggleBtn) {
    toggleBtn.classList.toggle('off', !leadsCreatePermissionEnabled);
  }
}

function toggleContactsCreatePermission(showNotice = false) {
  contactsCreatePermissionEnabled = !contactsCreatePermissionEnabled;
  const btn = document.getElementById('btnPermContactAdd');
  const box = document.getElementById('boxPermContactAdd');
  if (btn) btn.classList.toggle('checked', contactsCreatePermissionEnabled);
  if (box) box.classList.toggle('checked', contactsCreatePermissionEnabled);

  if (showNotice) {
    showToast(
      contactsCreatePermissionEnabled ? 'Permission Granted' : 'Permission Revoked',
      `Contacts Create permission is ${contactsCreatePermissionEnabled ? 'ENABLED' : 'DISABLED'}.`,
      'info',
      2500
    );
  }
}

function saveContactsPermissions() {
  updateLeadsCreatePermissionUI();
}

// =========================================================================
// AC DEMO & TEST CONTROLS (Interactive Verification Suite)
// =========================================================================
function setupAcDemoControls() {
  const btnToggle = document.getElementById('btnToggleAcDemo');
  const panel = document.getElementById('acDemoPanel');
  const chevron = document.getElementById('acDemoBarChevron');

  if (btnToggle && panel) {
    btnToggle.addEventListener('click', () => {
      const isOpen = panel.classList.toggle('open');
      if (chevron) chevron.textContent = isOpen ? '▼' : '▲';
    });
  }

  // Quick Test AC-2.1 & AC-4.1: New Lead Flow ("Orion Freight")
  const btnTestNewParty = document.getElementById('btnTestAcNewParty');
  if (btnTestNewParty) {
    btnTestNewParty.addEventListener('click', () => {
      showContactsModule();
      openContactDrawer();

      // Pre-fill contact details
      const first = document.getElementById('contactFirstName');
      const last = document.getElementById('contactLastName');
      const email = document.getElementById('contactEmail');
      const party = document.getElementById('contactPartyInput');

      if (first) first.value = 'David';
      if (last) last.value = 'Lee';
      if (email) email.value = 'david.lee@orionfreight.com';
      if (party) {
        party.value = 'Orion Freight';
        currentTypedPartyValue = 'Orion Freight';
      }

      // Trigger Address Capture Modal directly with Example-b values
      openAddressCaptureModal('Orion Freight');

      const addrInput = document.getElementById('addrCaptureAddress');
      const cityInput = document.getElementById('addrCaptureCity');
      const countrySelect = document.getElementById('addrCaptureCountry');

      if (addrInput) addrInput.value = '12 Marina Blvd';
      if (cityInput) cityInput.value = 'Singapore';
      if (countrySelect) countrySelect.value = 'SG';

      showToast(
        'AC-2.1 Scenario Active',
        'Typed "Orion Freight" (unmatched). Address popup opened with Singapore location ready to Save.',
        'info',
        4500
      );
    });
  }

  // Quick Test AC-5.1: Existing Match ("  rgb exports  ")
  const btnTestExistingMatch = document.getElementById('btnTestAcExistingMatch');
  if (btnTestExistingMatch) {
    btnTestExistingMatch.addEventListener('click', () => {
      showContactsModule();
      openContactDrawer();

      const party = document.getElementById('contactPartyInput');
      if (party) {
        party.value = '  rgb exports  ';
        currentTypedPartyValue = '  rgb exports  ';
        party.focus();
        setTimeout(() => {
          party.blur();
          showToast(
            'AC-5.1 Verified',
            'Defocused "  rgb exports  ": Matched existing "RGB Exports (Dubai, UAE)". Popup did NOT open.',
            'success',
            4500
          );
        }, 100);
      }
    });
  }

  // Quick Test AC-5.2: Vendor Role Excluded ("Acme Corp")
  const btnTestVendor = document.getElementById('btnTestAcVendorExcluded');
  if (btnTestVendor) {
    btnTestVendor.addEventListener('click', () => {
      showContactsModule();
      openContactDrawer();

      const party = document.getElementById('contactPartyInput');
      if (party) {
        party.value = 'Acme Corp';
        currentTypedPartyValue = 'Acme Corp';
        handlePartyBlur();
        showToast(
          'AC-5.2 Verified',
          '"Acme Corp" exists as Vendor (excluded from lookup). Address capture opened to create new Lead!',
          'info',
          4500
        );
      }
    });
  }

  // Quick Test AC-4.2: Duplicate Email Validation
  const btnTestDupEmail = document.getElementById('btnTestAcDuplicateEmail');
  if (btnTestDupEmail) {
    btnTestDupEmail.addEventListener('click', () => {
      showContactsModule();
      openContactDrawer();

      const first = document.getElementById('contactFirstName');
      const last = document.getElementById('contactLastName');
      const email = document.getElementById('contactEmail');
      const party = document.getElementById('contactPartyInput');

      if (first) first.value = 'Test';
      if (last) last.value = 'User';
      if (email) email.value = 'abcd@rajlog.com'; // Existing email row 1

      // Set pending party
      pendingNewParty = {
        companyName: 'Apex Quantum Log',
        address: '100 Ocean Way',
        city: 'Dubai',
        countryCode: 'AE',
        countryName: 'United Arab Emirates'
      };
      if (party) party.value = 'Apex Quantum Log';
      const badge = document.getElementById('partyPendingBadge');
      if (badge) badge.style.display = 'block';

      setTimeout(() => {
        handleSaveContact();
      }, 200);
    });
  }

  // Toggle AC-5.6 Leads Create Permission
  const btnTogglePerm = document.getElementById('btnTogglePermLeadAdd');
  if (btnTogglePerm) {
    btnTogglePerm.addEventListener('click', () => {
      toggleLeadsCreatePermission(true);
    });
  }

  // Toggle AC-5.5 Network Timeout Simulation
  const btnToggleTimeout = document.getElementById('btnToggleNetworkTimeout');
  const timeoutStatusText = document.getElementById('networkTimeoutStatusText');
  if (btnToggleTimeout) {
    btnToggleTimeout.addEventListener('click', () => {
      simulatePartySearchTimeout = !simulatePartySearchTimeout;
      if (timeoutStatusText) {
        timeoutStatusText.textContent = `Party Search: [${simulatePartySearchTimeout ? 'TIMEOUT SIMULATED' : 'NORMAL'}]`;
      }
      btnToggleTimeout.classList.toggle('off', simulatePartySearchTimeout);
      showToast(
        simulatePartySearchTimeout ? 'Search Timeout Enabled' : 'Search Normal',
        simulatePartySearchTimeout
          ? 'AC-5.5: Non-blocking "Unable to search existing parties right now" warning will show when typing.'
          : 'Normal search connectivity restored.',
        simulatePartySearchTimeout ? 'warning' : 'success',
        3500
      );
    });
  }

  // Reset Prototype Data
  const btnReset = document.getElementById('btnResetContactsData');
  if (btnReset) {
    btnReset.addEventListener('click', () => {
      contactsData = JSON.parse(JSON.stringify(INITIAL_CONTACTS));
      leadsCreatePermissionEnabled = true;
      simulatePartySearchTimeout = false;
      updateLeadsCreatePermissionUI();
      if (timeoutStatusText) timeoutStatusText.textContent = 'Party Search: [NORMAL]';
      if (btnToggleTimeout) btnToggleTimeout.classList.remove('off');
      renderContactsTable();
      showToast('Prototype Reset', 'Contacts data and permissions reset to initial screenshot state.', 'info', 2500);
    });
  }
}


