// Test utilities
function generateSampleData() {
  return {
    happyRows: [
      // Happy Row 1 - Complete and valid data
      {
        "Load ID": "TL123456",
        "Pickup": "123 Main St, Atlanta, GA 30303",
        "PU Time": "2025-09-10T14:00:00Z",
        "Delivery": "456 Oak Ave, Miami, FL 33101",
        "DEL Time": "2025-09-11T16:00:00Z",
        "Status": "assigned",
        "Driver": "John Smith",
        "Phone": "404-555-0123",
        "Unit": "4721",
        "Customer": "BigShipper Inc"
      },
      // Happy Row 2 - Complete data with different format dates (will be normalized)
      {
        "Load ID": "TL789012",
        "Pickup": "789 Pine Rd, Chicago, IL 60601",
        "PU Time": "9/11/2025 09:30 AM ET",
        "Delivery": "321 Elm St, Dallas, TX 75201",
        "DEL Time": "9/12/2025 02:15 PM CT",
        "Status": "in_transit",
        "Driver": "Sarah Johnson",
        "Phone": "312-555-0456",
        "Unit": "5832",
        "Customer": "MegaFreight LLC"
      }
    ],
    brokenRows: [
      // Broken Row 1 - Missing required fields (broker/customer and addresses)
      {
        "Load ID": "TL345678",
        "PU Time": "2025-09-13T10:00:00Z",
        "DEL Time": "2025-09-14T11:00:00Z",
        "Status": "pending",
        "Driver": "Mike Wilson",
        "Phone": "555-0789",
        "Unit": "3914",
        "Customer": "", // Empty required field
        "Pickup": null, // Missing required field as null
        "Delivery": "   " // Empty string with spaces
      // Broken Row 2 - Invalid dates and duplicate Load ID
      {
        "Load ID": "TL123456", // Duplicate of happy row 1
        "Pickup": "567 Beach Blvd, LA, CA 90001",
        "PU Time": "Invalid Date",
        "Delivery": "890 Mountain View, Denver, CO 80201",
        "DEL Time": "25/13/2025", // Invalid date format
        "Status": "ASSIGNED", // Inconsistent status capitalization
        "Driver": "Lisa Brown",
        "Phone": "213-555-0321",
        "Unit": "2947",
        "Customer": "FastFreight Co"
      },
      // Broken Row 3 - Partial/incomplete addresses
      {
        "Load ID": "TL901234",
        "Pickup": "Houston", // Incomplete address
        "PU Time": "2025-09-15T13:00:00Z",
        "Delivery": "TX", // Incomplete address
        "DEL Time": "2025-09-16T15:00:00Z",
        "Status": "assigned",
        "Driver": "Tom Davis",
        "Unit": "6103",
        "Customer": "QuickShip Inc"
      }
    ]
  };
}

function runTests() {
  const tests = [
    testDateTimeNormalization,
    testAddressValidation,
    testLoadIdValidation
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(test => {
    try {
      test();
      passed++;
      Logger.log(`✅ ${test.name} passed`);
    } catch (e) {
      failed++;
      Logger.log(`❌ ${test.name} failed: ${e.message}`);
    }
  });
  
  Logger.log(`\nTest Summary: ${passed} passed, ${failed} failed`);
}

// Pure function tests
function testDateTimeNormalization() {
  // Test cases for datetime normalization
  const cases = [
    {
      input: "9/10/2025 2:30 PM ET",
      expected: "2025-09-10T18:30:00Z" // ET + 4 = UTC
    },
    {
      input: "2025-09-10 14:00:00",
      expected: "2025-09-10T18:00:00Z"
    },
    {
      input: "10-Sep-2025 09:15",
      expected: "2025-09-10T13:15:00Z"
    },
    {
      input: "Invalid Date",
      expected: null
    }
  ];
  
  cases.forEach(({input, expected}, i) => {
    const result = normalizeDateTime(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} failed: expected ${expected}, got ${result}`);
    }
  });
}

function testAddressValidation() {
  const cases = [
    {
      input: "123 Main St, Atlanta, GA 30303",
      expected: true
    },
    {
      input: "Houston", // Too short/incomplete
      expected: false
    },
    {
      input: "", // Empty
      expected: false
    }
  ];
  
  cases.forEach(({input, expected}, i) => {
    const result = validateAddress(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} failed: expected ${expected}, got ${result}`);
    }
  });
}



function testLoadIdValidation() {
  const existingIds = ["TL123456", "TL789012"];
  
  const cases = [
    {
      input: "TL345678",
      expected: true // New ID
    },
    {
      input: "TL123456",
      expected: false // Duplicate
    },
    {
      input: "",
      expected: false // Empty
    }
  ];
  
  cases.forEach(({input, expected}, i) => {
    const result = validateLoadId(input, existingIds);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} failed: expected ${expected}, got ${result}`);
    }
  });
}

// Pure utility functions being tested
function normalizeDateTime(input) {
  if (!input) return null;
  
  try {
    const date = new Date(input);
    if (isNaN(date.getTime())) return null;
    
    // Assume ET if no timezone specified
    if (!input.includes('Z') && !input.includes('+') && !input.match(/[A-Z]{2,3}$/)) {
      date.setHours(date.getHours() + 4); // ET -> UTC
    }
    
    return date.toISOString();
  } catch {
    return null;
  }
}

function validateAddress(address) {
  if (!address) return false;
  
  // Basic validation: should have some minimum length and contain a comma
  // You can make this more sophisticated based on your needs
  return address.length >= 10 && address.includes(',');
}


function validateLoadId(loadId, existingIds) {
  if (!loadId) return false;
  return !existingIds.includes(loadId);
}

// New test functions for empty/missing values
function testRequiredFieldValidation() {
  const requiredFields = REQUIRED_FIELDS; // Using the constant from Code.gs
  const cases = [
    {
      input: {
        loadId: "TL123",
        fromAddress: "123 Main St",
        fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z",
        toAddress: "456 Oak Ave",
        toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z",
        status: "assigned",
        driverName: "John Smith",
        unitNumber: "4721",
        broker: "BigShipper Inc"
      },
      expected: true // All required fields present
    },
    {
      input: {
        loadId: "TL123",
        // Missing fromAddress
        fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z",
        toAddress: "456 Oak Ave",
        toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z",
        status: "assigned",
        driverName: "John Smith",
        unitNumber: "4721",
        broker: "BigShipper Inc"
      },
      expected: false // Missing required field
    },
    {
      input: {
        loadId: "",  // Empty required field
        fromAddress: "123 Main St",
        fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z",
        toAddress: "456 Oak Ave",
        toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z",
        status: "assigned",
        driverName: "John Smith",
        unitNumber: "4721",
        broker: "BigShipper Inc"
      },
      expected: false // Empty required field
    }
  ];
  
  cases.forEach(({input, expected}, i) => {
    const result = validateRequiredFields(input, requiredFields);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} failed: expected ${expected}, got ${result}`);
    }
  });
}

function testEmptyCellDetection() {
  const cases = [
    {
      input: "",
      expected: true
    },
    {
      input: "   ",
      expected: true
    },
    {
      input: null,
      expected: true
    },
    {
      input: undefined,
      expected: true
    },
    {
      input: "Not empty",
      expected: false
    }
  ];
  
  cases.forEach(({input, expected}, i) => {
    const result = isEmptyCell(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} failed: expected ${expected}, got ${result}`);
    }
  });
}

// Pure utility functions for empty/missing value validation
function validateRequiredFields(data, requiredFields) {
  return requiredFields.every(field => {
    const value = data[field];
    return value !== undefined && value !== null && value.toString().trim() !== '';
  });
}

function isEmptyCell(value) {
  if (value === null || value === undefined) return true;
  return value.toString().trim() === '';
}

function insertSampleData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sampleData = generateSampleData();
  
  // Clear existing data
  sheet.clear();
  
  // Get all unique headers
  const allHeaders = new Set();
  [...sampleData.happyRows, ...sampleData.brokenRows].forEach(row => {
    Object.keys(row).forEach(header => allHeaders.add(header));
  });
  
  // Insert headers
  const headers = Array.from(allHeaders);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Insert happy rows
  const happyRowsValues = sampleData.happyRows.map(row => 
    headers.map(header => row[header] || '')
  );
  if (happyRowsValues.length > 0) {
    sheet.getRange(2, 1, happyRowsValues.length, headers.length)
      .setValues(happyRowsValues);
  }
  
  // Insert broken rows
  const brokenRowsValues = sampleData.brokenRows.map(row => 
    headers.map(header => row[header] || '')
  );
  if (brokenRowsValues.length > 0) {
    sheet.getRange(2 + happyRowsValues.length, 1, brokenRowsValues.length, headers.length)
      .setValues(brokenRowsValues);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}
