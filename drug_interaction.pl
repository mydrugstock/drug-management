% ==========================================
% Drug Interaction Facts
% ==========================================

% Category X
interacts(amiloride_hydrochlorothiazide, spironolactone, x, major,
'Concurrent use of Potassium-Sparing Diuretics increases the risk of severe hyperkalemia.').

% Category D
interacts(atenolol, methyldopa, d, moderate,
'Risk of rebound hypertension and AV block.').

interacts(captopril, losartan, d, moderate,
'Dual RAS blockade increases the risk of renal impairment, hypotension, and hyperkalemia.').

interacts(enalapril, losartan, d, major,
'Dual RAS blockade increases the risk of renal impairment, hypotension, and hyperkalemia.').

interacts(carvedilol, methyldopa, d, moderate,
'Risk of rebound hypertension, bradycardia, and AV block.').

interacts(metoprolol, methyldopa, d, moderate,
'Risk of rebound hypertension and sinus node dysfunction.').

interacts(pioglitazone, glipizide, d, moderate,
'Increased risk of hypoglycemia.').

interacts(propranolol, methyldopa, d, moderate,
'Risk of rebound hypertension and sinus node dysfunction.').

% Category C
interacts(pioglitazone, insulin_nph, c, major,
'Increased risk of hypoglycemia and exacerbation of heart failure.').

% Dose Limitation
interacts(simvastatin, amlodipine, dose_limit, moderate,
'Limit simvastatin dose to 20 mg/day to reduce the risk of myopathy and rhabdomyolysis.').

% ==========================================
% Drug Interaction Rules
% ==========================================

check_interaction(DrugA, DrugB, Risk, Severity, Warning) :-
    interacts(DrugA, DrugB, Risk, Severity, Warning).

check_interaction(DrugA, DrugB, Risk, Severity, Warning) :-
    interacts(DrugB, DrugA, Risk, Severity, Warning).