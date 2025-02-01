// sheetConfig.js
module.exports = {
    // Default configuration for a sheet that must have Name, Amount, Date, Verified columns
    default: {
      // The expected Excel column headers (keys) are mapped to database field names (values)
      columnMapping: {
        Name: 'name',
        Amount: 'amount',
        Date: 'date',
        Verified: 'verified'
      },
      // Validation rules for each column.
      // For future extensions, you could add configurations for specific sheet names.
      validationRules: {
        Name: {
          required: true,
          type: 'string'
        },
        Amount: {
          required: true,
          type: 'number',
          min: 0.01 // must be greater than zero
        },
        Date: {
          required: true,
          type: 'date',
          currentMonth: true // date must fall within the current month
        },
        Verified: {
          required: false,
          allowedValues: ['Yes', 'No']
        }
      }
    }
    // You can add additional sheet configurations here by key (for example, "Invoices", etc.)
  };
  