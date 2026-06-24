/**
 * myDoc - Configuration & Constants
 */

// --- Configuration ---
// REPLACE WITH YOUR SPREADSHEET ID (found in the spreadsheet URL)
// Example: https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit
const SPREADSHEET_ID = '1him9YU3PoWMIUc21NmN7xJgFisJ-_FAZ8RTkFvKGN2Q'; // Updated from provided URL

// REPLACE WITH YOUR GOOGLE CALENDAR ID (e.g., 'primary' or 'c_xxxxxxxx@group.calendar.google.com')
const CALENDAR_ID = 'c_b0a73eb7d9a6d0cdcc2f1e6d021dff719917e8d35f571d9855c5910e91bbc4f9@group.calendar.google.com';

// --- Authorized Users (Hardcoded for Testing/Debugging) ---
// Add your authorized users here with their roles
// Format: { email: 'user@example.com', name: 'User Name', role: 'ROLE', status: 'Active' }
const AUTHORIZED_USERS = [
  { email: 'hudson.eboso@techbrain.africa', name: 'Hudson Eboso', role: 'ADMIN', status: 'Active' },
  { email: 'ebosoh@gmail.com', name: 'Tajiri', role: 'ADMIN', status: 'Active' },
  // Add more users below:
  // { email: 'doctor@example.com', name: 'Dr. Smith', role: 'DOCTOR', status: 'Active' },
  // { email: 'nurse@example.com', name: 'Nurse Jane', role: 'NURSE', status: 'Active' },
];

// --- Database Sheet Configuration ---
const SHEET_CONFIG = {
  'Registration_Records': [
    'patient_id', 'registration_date', 'national_id', 'NHIF_no', 'first_name', 'middle_name',
    'last_name', 'date_of_birth', 'gender', 'marital_status', 'phone_number', 'email',
    'county', 'sub_county', 'ward', 'language', 'next_of_kin', 'NOK_relationship',
    'NOK_phone', 'blood_group', 'known_allergies', 'is_active', 'registered_by_staff_ID'
  ],
  'Triage_Records': [
    'encounter_id', 'patient_id', 'triage_date_time', 'chief_complain', 'complain_duration',
    'bp_systolic', 'bp_diastolic', 'heart_rate_bpm', 'resp_rate_breaths', 'temp_c',
    'oxygen_saturation', 'weight', 'height', 'BMI', 'triage_priority', 'triage_nurse_id'
  ],
  'Treatment_Records': [
    'treatment_id', 'encounter_id', 'patient_id', 'doctor_id', 'visit_date_time',
    'history_present_illness', 'physical_exam_findings', 'diagnosis_primary_code',
    'diagnosis_primary_desc', 'diagnosis_secondary_code', 'prescribed_medications',
    'labd_referral_ids', 'treatment_plan', 'follow_up_date', 'disposition',
    'consultation_stage', 'lab_results_reviewed', 'lab_review_date', 'initial_consultation_date'
  ],
  'Laboratory_Orders': [
    'lab_order_id', 'patient_id', 'encounter_id', 'ordering_doctor_id', 'order_date_time',
    'test_name', 'test_code', 'test_status', 'result_value', 'result_units',
    'reference_range', 'abnormal_flag', 'lab_technologist_id', 'detailed_description', 'file_url'
  ],
  'Laboratory_Inventory_Records': [
    'date', 'item_id', 'item_name', 'transaction_type', 'quantity_in', 'unit_cost',
    'quantity_out', 'stock_balance'
  ],
  'Billing_and_Payment_Records': [
    'payment_transaction_id', 'encounter_id', 'patient_id', 'billing_date_time',
    'service_category', 'total_charge_amount', 'amount_paid_today', 'payment_method',
    'transaction_reference', 'outstanding_balance', 'payment_received_by_staff_id', 'notes'
  ],
  'Debtors_Records': [
    'debtor_id', 'patient_id', 'patient_name', 'phone_number', 'encounter_id',
    'service_date', 'service_description', 'total_amount', 'amount_paid',
    'outstanding_balance', 'last_payment_date', 'status', 'notes'
  ],
  'FP_Client_Records': [
    'client_number', 'registration_date', 'client_name', 'client_type', 'first_ever_user',
    'age', 'sex', 'disability_status', 'telephone', 'county', 'subcounty',
    'village_estate', 'registered_by'
  ],
  'FP_Tracking_Records': [
    'tracking_id', 'client_number', 'client_name', 'phone_number', 'fp_type',
    'dispense_date', 'dosage', 'duration_days', 'return_date', 'notes',
    'status', 'dispensed_by'
  ],
  'FP_Archives': [
    'tracking_id', 'client_number', 'client_name', 'phone_number', 'fp_type',
    'dispense_date', 'dosage', 'duration_days', 'return_date', 'notes',
    'status', 'dispensed_by', 'archived_date', 'archived_by'
  ],
  'ICD11_Codes': [
    'code', 'description', 'category', 'common_name'
  ]
};
