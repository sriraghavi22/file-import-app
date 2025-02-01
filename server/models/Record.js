// models/Record.js
const mongoose = require('mongoose');

const RecordSchema = new mongoose.Schema({
  name: { type: String, required: true },
  amount: { type: Number, required: true },
  date: { type: Date, required: true },
  verified: { type: String, enum: ['Yes', 'No'], default: 'No' }
  // You can extend with additional fields for different sheets
});

module.exports = mongoose.model('Record', RecordSchema);
