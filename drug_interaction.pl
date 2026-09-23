:- encoding(utf8).
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


% ==========================================================
% LAB REFERENCE RANGES & STATUS (สูง / ต่ำ / ปกติ)
% ==========================================================
% ⚠️ ค่าปกติ/threshold ต่อไปนี้อ้างอิงหลักการทั่วไปเพื่อความปลอดภัย
% เบื้องต้นเท่านั้น ไม่ใช่คำวินิจฉัยทางการแพทย์ ทีมเภสัชกร/แพทย์ที่
% ปรึกษาโปรเจกต์ต้องตรวจทาน threshold ทั้งหมดก่อนใช้งานกับผู้ป่วยจริง
% ==========================================================

% lab_range(LabKey, NormalLow, NormalHigh, Unit).
lab_range(egfr, 90, 999, 'ml/min/1.73m2').
lab_range(potassium, 3.5, 5.0, 'mEq/L').
lab_range(alt, 0, 40, 'U/L').

% lab_status(LabKey, Value, Status, Description)
% Status ที่เป็นไปได้: normal / low / high
% ==== eGFR (การทำงานของไต) ====
lab_status(egfr, Value, normal, 'การทำงานของไตปกติ') :-
    number(Value), Value >= 90, !.
lab_status(egfr, Value, low, 'การทำงานของไตลดลงเล็กน้อย (CKD stage 2)') :-
    number(Value), Value >= 60, Value < 90, !.
lab_status(egfr, Value, low, 'การทำงานของไตลดลงปานกลาง (CKD stage 3)') :-
    number(Value), Value >= 30, Value < 60, !.
lab_status(egfr, Value, low, 'การทำงานของไตลดลงรุนแรง (CKD stage 4)') :-
    number(Value), Value >= 15, Value < 30, !.
lab_status(egfr, Value, low, 'ไตวาย (CKD stage 5 / Kidney failure)') :-
    number(Value), Value < 15, !.

% ==== Potassium (K+) ====
lab_status(potassium, Value, high, 'โพแทสเซียมสูงรุนแรง (Severe hyperkalemia)') :-
    number(Value), Value > 6.0, !.
lab_status(potassium, Value, high, 'โพแทสเซียมสูงปานกลาง (Moderate hyperkalemia)') :-
    number(Value), Value > 5.5, Value =< 6.0, !.
lab_status(potassium, Value, high, 'โพแทสเซียมสูงเล็กน้อย (Mild hyperkalemia)') :-
    number(Value), Value > 5.0, Value =< 5.5, !.
lab_status(potassium, Value, normal, 'โพแทสเซียมปกติ') :-
    number(Value), Value >= 3.5, Value =< 5.0, !.
lab_status(potassium, Value, low, 'โพแทสเซียมต่ำเล็กน้อย (Mild hypokalemia)') :-
    number(Value), Value >= 3.0, Value < 3.5, !.
lab_status(potassium, Value, low, 'โพแทสเซียมต่ำรุนแรง (Severe hypokalemia)') :-
    number(Value), Value < 3.0, !.

% ==== ALT (การทำงานของตับ) ====
lab_status(alt, Value, high, 'ALT สูงมาก (>5 เท่าของค่าปกติ)') :-
    number(Value), Value > 200, !.
lab_status(alt, Value, high, 'ALT สูงปานกลาง (3-5 เท่าของค่าปกติ)') :-
    number(Value), Value > 120, Value =< 200, !.
lab_status(alt, Value, high, 'ALT สูงเล็กน้อย (1-3 เท่าของค่าปกติ)') :-
    number(Value), Value > 40, Value =< 120, !.
lab_status(alt, Value, normal, 'ALT ปกติ') :-
    number(Value), Value >= 0, Value =< 40, !.


% ==========================================================
% LAB-BASED DOSE ADJUSTMENT (ยา + ค่า lab -> คำแนะนำ)
% ==========================================================
% lab_adjustment(DrugAtom, LabKey, MinValue, MaxValue, Severity, Recommendation).
% ช่วง [MinValue, MaxValue] เป็นแบบ inclusive ทั้งสองด้าน

% ---- eGFR ----
lab_adjustment(glipizide, egfr, 0, 30, caution,
'Glipizide ไม่ถูกขับผ่านไตเป็นหลักและไม่มี active metabolite จึงค่อนข้างปลอดภัยกว่าซัลโฟนิลยูเรียตัวอื่นในผู้ป่วยไตเสื่อมรุนแรง แต่ยังควรเริ่มขนาดต่ำและติดตามภาวะน้ำตาลต่ำใกล้ชิด (eGFR < 30)').

lab_adjustment(insulin_nph, egfr, 0, 30, caution,
'ความต้องการอินซูลินมักลดลงเมื่อไตเสื่อม ควรพิจารณาลดขนาดและติดตามน้ำตาลถี่ขึ้น (eGFR < 30) — ปรับขนาดตามการตอบสนอง ไม่ใช่เปลี่ยนยา').

lab_adjustment(spironolactone, egfr, 0, 30, contraindicated,
'ห้ามใช้ (contraindicated) เสี่ยง hyperkalemia รุนแรง (eGFR < 30 ตาม product labeling)').

lab_adjustment(spironolactone, egfr, 30, 50, caution,
'ใช้ด้วยความระมัดระวังสูง ติดตามค่า K+ ใกล้ชิด (eGFR 30-49)').

lab_adjustment(amiloride_hydrochlorothiazide, egfr, 0, 30, caution,
'ยากลุ่ม potassium-sparing diuretic ร่วมกับ thiazide ความเสี่ยง hyperkalemia สูงขึ้นมากเมื่อไตเสื่อม ควรเลี่ยง (eGFR < 30)').

lab_adjustment(enalapril, egfr, 0, 30, caution,
'โดยทั่วไปแนะนำ คงการใช้ยาต่อพร้อมติดตามใกล้ชิด มากกว่าเปลี่ยนยา — เริ่มขนาดต่ำ ติดตาม K+/Creatinine หลังเริ่ม/ปรับยา 2-3 สัปดาห์ (eGFR < 30)').

lab_adjustment(captopril, egfr, 0, 30, caution,
'โดยทั่วไปแนะนำ คงการใช้ยาต่อพร้อมติดตามใกล้ชิด มากกว่าเปลี่ยนยา — เริ่มขนาดต่ำ ติดตาม K+/Creatinine หลังเริ่ม/ปรับยา 2-3 สัปดาห์ (eGFR < 30)').

lab_adjustment(losartan, egfr, 0, 30, caution,
'โดยทั่วไปแนะนำ คงการใช้ยาต่อพร้อมติดตามใกล้ชิด มากกว่าเปลี่ยนยา — เริ่มขนาดต่ำ ติดตาม K+/Creatinine หลังเริ่ม/ปรับยา 2-3 สัปดาห์ (eGFR < 30)').

% ---- Potassium (K+) ----
lab_adjustment(spironolactone, potassium, 5.5, 999, contraindicated,
'K+ สูง (>5.5) ร่วมกับยากลุ่ม potassium-sparing diuretic — เสี่ยง hyperkalemia รุนแรง ควรพิจารณาหยุดยาและติดตาม K+ ซ้ำ ก่อนตัดสินใจให้ยาต่อ').

lab_adjustment(amiloride_hydrochlorothiazide, potassium, 5.5, 999, contraindicated,
'K+ สูง (>5.5) ร่วมกับยากลุ่ม potassium-sparing diuretic — เสี่ยง hyperkalemia รุนแรง ควรพิจารณาหยุดยาและติดตาม K+ ซ้ำ ก่อนตัดสินใจให้ยาต่อ').

lab_adjustment(enalapril, potassium, 5.5, 6.0, caution,
'K+ สูงระดับปานกลาง (5.5-6.0) ร่วมกับ ACEi — ควรติดตาม K+ ใกล้ชิด พิจารณาลดขนาดหรือหยุดยาตามดุลยพินิจแพทย์').
lab_adjustment(enalapril, potassium, 6.0, 999, danger,
'K+ สูงรุนแรง (>6.0) ร่วมกับ ACEi — เสี่ยง cardiac arrhythmia ควรพิจารณาหยุดยาและส่งตรวจ EKG/แก้ไข K+ เร่งด่วนตามดุลยพินิจแพทย์').

lab_adjustment(captopril, potassium, 5.5, 6.0, caution,
'K+ สูงระดับปานกลาง (5.5-6.0) ร่วมกับ ACEi — ควรติดตาม K+ ใกล้ชิด พิจารณาลดขนาดหรือหยุดยาตามดุลยพินิจแพทย์').
lab_adjustment(captopril, potassium, 6.0, 999, danger,
'K+ สูงรุนแรง (>6.0) ร่วมกับ ACEi — เสี่ยง cardiac arrhythmia ควรพิจารณาหยุดยาและส่งตรวจ EKG/แก้ไข K+ เร่งด่วนตามดุลยพินิจแพทย์').

lab_adjustment(losartan, potassium, 5.5, 6.0, caution,
'K+ สูงระดับปานกลาง (5.5-6.0) ร่วมกับ ARB — ควรติดตาม K+ ใกล้ชิด พิจารณาลดขนาดหรือหยุดยาตามดุลยพินิจแพทย์').
lab_adjustment(losartan, potassium, 6.0, 999, danger,
'K+ สูงรุนแรง (>6.0) ร่วมกับ ARB — เสี่ยง cardiac arrhythmia ควรพิจารณาหยุดยาและส่งตรวจ EKG/แก้ไข K+ เร่งด่วนตามดุลยพินิจแพทย์').

% ---- ALT ----
lab_adjustment(simvastatin, alt, 120, 999, caution,
'ALT สูงเกิน ~3 เท่าของค่าปกติทั่วไป (สมมติ ULN ~40 U/L) ร่วมกับการใช้ statin — ตาม product labeling ทั่วไปแนะนำพิจารณาหยุดหรือลดขนาดยา และตรวจ LFT ซ้ำก่อนตัดสินใจ ไม่ควรฟันธงเปลี่ยนยาโดยไม่มีแพทย์พิจารณา').

lab_adjustment(pioglitazone, alt, 120, 999, caution,
'ALT สูงเกิน ~3 เท่าของค่าปกติทั่วไป — Pioglitazone มีคำเตือนเรื่อง hepatotoxicity ตาม labeling ควรหลีกเลี่ยงจนกว่า LFT จะกลับสู่ระดับที่ยอมรับได้ และให้แพทย์ประเมินสาเหตุ liver enzyme สูงก่อน').

% check_lab_adjustment/5: หาคำแนะนำปรับยา (ถ้ามี) สำหรับ Drug + Lab + Value ที่กำหนด
check_lab_adjustment(Drug, Lab, Value, Severity, Recommendation) :-
    lab_adjustment(Drug, Lab, MinValue, MaxValue, Severity, Recommendation),
    number(Value),
    Value >= MinValue,
    Value =< MaxValue.
