# ============================================================
# CREATE DRUG INTERACTION CONSULT PDF
# ============================================================

@app.route(
    "/create-consult-pdf",
    methods=["POST"]
)
def create_consult_pdf():

    try:

        # ====================================================
        # รับข้อมูลผู้ป่วย
        # ====================================================

        hn = request.form.get(
            "hn",
            ""
        )

        patient_name = request.form.get(
            "patient_name",
            ""
        )

        age = request.form.get(
            "age",
            ""
        )

        ward = request.form.get(
            "ward",
            "ไม่ได้ระบุ"
        )

        diagnosis = request.form.get(
            "diagnosis",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        weight = request.form.get(
            "weight",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        height = request.form.get(
            "height",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        lab = request.form.get(
            "lab",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        renal_function = request.form.get(
            "renal_function",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        hepatic_function = request.form.get(
            "hepatic_function",
            "ไม่ได้ระบุในข้อมูลใบสั่งยา"
        )

        dispense_date = request.form.get(
            "dispense_date",
            ""
        )

        appointment_date = request.form.get(
            "appointment_date",
            ""
        )


        # ====================================================
        # จำนวน Interaction
        # ====================================================

        try:

            interaction_count = int(
                request.form.get(
                    "interaction_count",
                    0
                )
            )

        except Exception:

            interaction_count = 0


        # ====================================================
        # รับ Interaction
        # ====================================================

        interactions = []


        for i in range(
            interaction_count
        ):

            interaction = {

                "Drug_1":
                    request.form.get(
                        "drug1_{}".format(i),
                        ""
                    ),

                "Drug_2":
                    request.form.get(
                        "drug2_{}".format(i),
                        ""
                    ),

                "Risk":
                    request.form.get(
                        "risk_{}".format(i),
                        ""
                    ),

                "Severity":
                    request.form.get(
                        "severity_{}".format(i),
                        ""
                    ),

                "Summary":
                    request.form.get(
                        "summary_{}".format(i),
                        ""
                    ),

                "Management":
                    request.form.get(
                        "management_{}".format(i),
                        ""
                    ),

                "Reference":
                    request.form.get(
                        "reference_{}".format(i),
                        ""
                    )
            }

            interactions.append(
                interaction
            )


        # ====================================================
        # ถ้าไม่มี Interaction
        # ====================================================

        if len(interactions) == 0:

            return """
            <h2>ไม่พบข้อมูล Drug Interaction</h2>

            <p>
            ไม่สามารถสร้าง Drug Interaction Consult
            เนื่องจากไม่พบข้อมูล Interaction
            ในฐานความรู้ของระบบ
            </p>

            <a href="/prescription">
                ← กลับไปตรวจใบสั่งยา
            </a>
            """


        # ====================================================
        # รับรายการยาปัจจุบันจาก Form
        # ====================================================

        try:

            medication_count = int(
                request.form.get(
                    "medication_count",
                    0
                )
            )

        except Exception:

            medication_count = 0


        medications = []


        for i in range(
            medication_count
        ):

            medications.append({

                "name":
                    request.form.get(
                        "med_name_{}".format(i),
                        ""
                    ),

                "strength":
                    request.form.get(
                        "med_strength_{}".format(i),
                        ""
                    ),

                "frequency":
                    request.form.get(
                        "med_frequency_{}".format(i),
                        ""
                    ),

                "route":
                    request.form.get(
                        "med_route_{}".format(i),
                        "ไม่ได้ระบุ"
                    ),

                "start_date":
                    request.form.get(
                        "med_start_date_{}".format(i),
                        "ไม่ได้ระบุ"
                    ),

                "quantity":
                    request.form.get(
                        "med_quantity_{}".format(i),
                        ""
                    )
            })


        # ====================================================
        # ถ้าไม่ได้ส่งรายการยา
        # ====================================================

        if len(medications) == 0:

            # สร้างรายการจากคู่ยาที่พบ Interaction
            # เพื่อไม่ให้ PDF ว่าง

            added_drugs = []

            for item in interactions:

                drug1 = item.get(
                    "Drug_1",
                    ""
                )

                drug2 = item.get(
                    "Drug_2",
                    ""
                )

                if drug1 and drug1 not in added_drugs:

                    medications.append({

                        "name": drug1,

                        "strength": "ไม่ได้ระบุ",

                        "frequency": "ไม่ได้ระบุ",

                        "route": "ไม่ได้ระบุ",

                        "start_date": "ไม่ได้ระบุ",

                        "quantity": "ไม่ได้ระบุ"
                    })

                    added_drugs.append(
                        drug1
                    )


                if drug2 and drug2 not in added_drugs:

                    medications.append({

                        "name": drug2,

                        "strength": "ไม่ได้ระบุ",

                        "frequency": "ไม่ได้ระบุ",

                        "route": "ไม่ได้ระบุ",

                        "start_date": "ไม่ได้ระบุ",

                        "quantity": "ไม่ได้ระบุ"
                    })

                    added_drugs.append(
                        drug2
                    )


        # ====================================================
        # สร้างชื่อไฟล์
        # ====================================================

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )


        safe_hn = str(
            hn
        ).replace(
            " ",
            "_"
        )


        filename = (
            "Drug_Interaction_Consult_HN_{}_{}.pdf"
            .format(
                safe_hn,
                timestamp
            )
        )


        pdf_path = os.path.join(
            CONSULT_FOLDER,
            filename
        )


        # ====================================================
        # Font
        # ====================================================

        normal_font, bold_font = (
            register_thai_fonts()
        )


        # ====================================================
        # สร้าง PDF
        # ====================================================

        pdf = canvas.Canvas(
            pdf_path,
            pagesize=A4
        )

        width, height = A4


        # ====================================================
        # Helper สำหรับขึ้นหน้าใหม่
        # ====================================================

        def check_page_space(
            current_y,
            required_space=80
        ):

            if current_y < required_space:

                pdf.showPage()

                return height - 50

            return current_y


        # ====================================================
        # HEADER
        # ====================================================

        pdf.setFont(
            bold_font,
            22
        )

        pdf.drawCentredString(
            width / 2,
            height - 45,
            "DRUG INTERACTION CONSULT"
        )


        pdf.setFont(
            bold_font,
            17
        )

        pdf.drawCentredString(
            width / 2,
            height - 72,
            "บันทึกปรึกษาปัญหายาระหว่างยาตีกัน"
        )


        y = height - 110


        # ====================================================
        # 1. SUBJECT / CONSULT REASON
        # ====================================================

        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "1. หัวข้อการปรึกษา (Subject / Consult Reason)"
        )

        y -= 28


        pdf.setFont(
            normal_font,
            15
        )


        subject_text = (
            "ขอปรึกษาเรื่อง Drug Interaction "
            "ในผู้ป่วย {} (HN: {}) "
            "อายุ {} ปี "
            "พบความเป็นไปได้ของปฏิกิริยาระหว่างยา"
        ).format(
            patient_name if patient_name else "ไม่ระบุชื่อ",
            hn if hn else "ไม่ระบุ",
            age if age else "ไม่ระบุ"
        )


        y = draw_wrapped_text(
            pdf,
            subject_text,
            60,
            y,
            max_chars=75,
            font=normal_font,
            size=15,
            line_height=21
        )


        y -= 8


        pdf.drawString(
            60,
            y,
            "วอร์ด/แผนก: {}".format(
                ward
            )
        )


        y -= 35


        # ====================================================
        # 2. CLINICAL CONTEXT
        # ====================================================

        y = check_page_space(
            y,
            150
        )


        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "2. ข้อมูลทางคลินิกที่สำคัญ (Clinical Context)"
        )

        y -= 28


        pdf.setFont(
            normal_font,
            15
        )


        clinical_items = [

            (
                "Diagnosis หลัก",
                diagnosis
            ),

            (
                "น้ำหนัก",
                weight
            ),

            (
                "ส่วนสูง",
                height
            ),

            (
                "ค่าทางห้องปฏิบัติการที่เกี่ยวข้อง",
                lab
            ),

            (
                "การทำงานของไต (CrCl / eGFR)",
                renal_function
            ),

            (
                "การทำงานของตับ",
                hepatic_function
            ),

            (
                "วันที่จ่ายยา",
                dispense_date
                if dispense_date
                else "ไม่ได้ระบุ"
            ),

            (
                "วันนัด",
                appointment_date
                if appointment_date
                else "ไม่ได้ระบุ"
            )
        ]


        for label, value in clinical_items:

            y = check_page_space(
                y,
                50
            )

            y = draw_wrapped_text(
                pdf,
                "{}: {}".format(
                    label,
                    value
                ),
                60,
                y,
                max_chars=75,
                font=normal_font,
                size=15,
                line_height=20
            )


        y -= 15


        # ====================================================
        # 3. CURRENT MEDICATION PROFILE
        # ====================================================

        y = check_page_space(
            y,
            180
        )


        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "3. รายการยาปัจจุบันของผู้ป่วย"
        )

        y -= 25


        pdf.setFont(
            normal_font,
            13
        )


        # Header ตาราง

        pdf.drawString(
            50,
            y,
            "ยา"
        )

        pdf.drawString(
            180,
            y,
            "ขนาดยา"
        )

        pdf.drawString(
            270,
            y,
            "วิธีใช้"
        )

        pdf.drawString(
            370,
            y,
            "Route"
        )

        pdf.drawString(
            440,
            y,
            "วันที่เริ่มยา"
        )


        y -= 8


        pdf.line(
            50,
            y,
            width - 50,
            y
        )


        y -= 20


        for medicine in medications:

            y = check_page_space(
                y,
                60
            )


            drug_name = str(
                medicine.get(
                    "name",
                    ""
                )
            )

            strength = str(
                medicine.get(
                    "strength",
                    "ไม่ได้ระบุ"
                )
            )

            frequency = str(
                medicine.get(
                    "frequency",
                    "ไม่ได้ระบุ"
                )
            )

            route = str(
                medicine.get(
                    "route",
                    "ไม่ได้ระบุ"
                )
            )

            start_date = str(
                medicine.get(
                    "start_date",
                    "ไม่ได้ระบุ"
                )
            )


            # จำกัดความยาวเพื่อไม่ให้ชนกัน

            if len(drug_name) > 20:

                drug_name = (
                    drug_name[:20]
                    + "..."
                )


            if len(str(strength)) > 12:

                strength = (
                    str(strength)[:12]
                    + "..."
                )


            if len(str(frequency)) > 14:

                frequency = (
                    str(frequency)[:14]
                    + "..."
                )


            if len(str(route)) > 10:

                route = (
                    str(route)[:10]
                    + "..."
                )


            if len(str(start_date)) > 14:

                start_date = (
                    str(start_date)[:14]
                    + "..."
                )


            pdf.drawString(
                50,
                y,
                drug_name
            )

            pdf.drawString(
                180,
                y,
                strength
            )

            pdf.drawString(
                270,
                y,
                frequency
            )

            pdf.drawString(
                370,
                y,
                route
            )

            pdf.drawString(
                440,
                y,
                start_date
            )


            y -= 23


        y -= 15


        # ====================================================
        # 4. IDENTIFIED DRUG INTERACTION
        # ====================================================

        y = check_page_space(
            y,
            220
        )


        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "4. ปัญหาและปฏิกิริยาระหว่างยาที่พบ"
        )

        y -= 28


        pdf.setFont(
            normal_font,
            15
        )


        for index, item in enumerate(
            interactions
        ):

            y = check_page_space(
                y,
                180
            )


            # ------------------------------------------------
            # Interaction number
            # ------------------------------------------------

            pdf.setFont(
                bold_font,
                16
            )

            pdf.drawString(
                60,
                y,
                "Interaction {}".format(
                    index + 1
                )
            )

            y -= 23


            pdf.setFont(
                normal_font,
                15
            )


            # ------------------------------------------------
            # คู่ยา
            # ------------------------------------------------

            drug1 = item.get(
                "Drug_1",
                ""
            )

            drug2 = item.get(
                "Drug_2",
                ""
            )


            y = draw_wrapped_text(
                pdf,
                "คู่ยาที่พบ: {} + {}".format(
                    drug1,
                    drug2
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            # ------------------------------------------------
            # Object / Precipitant
            # ------------------------------------------------

            y = draw_wrapped_text(
                pdf,
                "Object Drug: {}".format(
                    drug1
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            y = draw_wrapped_text(
                pdf,
                "Precipitant Drug: {}".format(
                    drug2
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            # ------------------------------------------------
            # Severity
            # ------------------------------------------------

            y = draw_wrapped_text(
                pdf,
                "Severity: {}".format(
                    item.get(
                        "Severity",
                        "ไม่ได้ระบุ"
                    )
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            # ------------------------------------------------
            # Risk / Evidence
            # ------------------------------------------------

            y = draw_wrapped_text(
                pdf,
                "Evidence / Risk: {}".format(
                    item.get(
                        "Risk",
                        "ไม่ได้ระบุ"
                    )
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            # ------------------------------------------------
            # Mechanism
            # ------------------------------------------------

            y = draw_wrapped_text(
                pdf,
                "กลไก / Clinical Significance: {}".format(
                    item.get(
                        "Summary",
                        "ไม่ได้ระบุในฐานความรู้ของระบบ"
                    )
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            # ------------------------------------------------
            # Reference
            # ------------------------------------------------

            y = draw_wrapped_text(
                pdf,
                "Reference: {}".format(
                    item.get(
                        "Reference",
                        "ไม่ได้ระบุ"
                    )
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


            y -= 15


        # ====================================================
        # 5. RECOMMENDATIONS
        # ====================================================

        y = check_page_space(
            y,
            180
        )


        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "5. ข้อเสนอแนะทางเภสัชบำบัด (Recommendations)"
        )

        y -= 28


        pdf.setFont(
            normal_font,
            15
        )


        pdf.drawString(
            60,
            y,
            "ข้อเสนอแนะจากฐานความรู้ของระบบ:"
        )

        y -= 23


        for index, item in enumerate(
            interactions
        ):

            management = item.get(
                "Management",
                ""
            )


            y = draw_wrapped_text(
                pdf,
                "{}. {}".format(
                    index + 1,
                    management
                    if management
                    else "ไม่ได้ระบุ"
                ),
                70,
                y,
                max_chars=70,
                font=normal_font,
                size=15,
                line_height=20
            )


        y -= 15


        # ----------------------------------------------------
        # Monitoring
        # ----------------------------------------------------

        y = draw_wrapped_text(
            pdf,
            "Monitoring parameters: "
            "ควรพิจารณาติดตามอาการไม่พึงประสงค์ "
            "และผลตรวจทางห้องปฏิบัติการที่เกี่ยวข้อง "
            "ตามความเหมาะสมของผู้ป่วย",
            60,
            y,
            max_chars=75,
            font=normal_font,
            size=15,
            line_height=20
        )


        y -= 10


        y = draw_wrapped_text(
            pdf,
            "การปรับขนาดยา / การเปลี่ยนยา / "
            "การเว้นระยะเวลาการบริหารยา: "
            "ให้แพทย์และทีมสหวิชาชีพพิจารณาตามบริบททางคลินิก",
            60,
            y,
            max_chars=75,
            font=normal_font,
            size=15,
            line_height=20
        )


        # ====================================================
        # 6. CONCLUSION & SIGN-OFF
        # ====================================================

        y = check_page_space(
            y,
            190
        )


        pdf.setFont(
            bold_font,
            18
        )

        pdf.drawString(
            50,
            y,
            "6. สรุปและลงชื่อผู้ปรึกษา"
        )

        y -= 28


        pdf.setFont(
            normal_font,
            15
        )


        conclusion = (
            "สรุป: พบข้อมูล Drug Interaction "
            "จำนวน {} รายการ จากฐานความรู้ของระบบ "
            "จึงขอปรึกษาแพทย์เพื่อพิจารณาความเหมาะสม "
            "ของการใช้ยาร่วมกัน"
        ).format(
            len(interactions)
        )


        y = draw_wrapped_text(
            pdf,
            conclusion,
            60,
            y,
            max_chars=75,
            font=normal_font,
            size=15,
            line_height=21
        )


        y -= 30


        pdf.drawString(
            60,
            y,
            "ผู้ปรึกษา: __________________________________________"
        )


        y -= 28


        pdf.drawString(
            60,
            y,
            "ตำแหน่ง: ภก. / นศ.ภ. _________________________________"
        )


        y -= 28


        pdf.drawString(
            60,
            y,
            "วันที่: ______________________________________________"
        )


        y -= 28


        pdf.drawString(
            60,
            y,
            "ช่องทางติดต่อกลับ: ___________________________________"
        )


        # ====================================================
        # FOOTER
        # ====================================================

        pdf.setFont(
            normal_font,
            10
        )


        pdf.drawCentredString(
            width / 2,
            25,
            "Medication Management / Clinical Decision Support System"
        )


        # ====================================================
        # SAVE PDF
        # ====================================================

        pdf.save()


        # ====================================================
        # ตรวจไฟล์
        # ====================================================

        if not os.path.exists(
            pdf_path
        ):

            raise Exception(
                "ไม่พบไฟล์ PDF หลังจากสร้าง"
            )


        # ====================================================
        # แสดงหน้า Success
        # ====================================================

        return render_template(
            "consult_success.html",
            filename=filename
        )


    except Exception as e:

        print(
            "=" * 70
        )

        print(
            "ERROR CREATE DRUG INTERACTION CONSULT PDF"
        )

        print(
            repr(e)
        )

        print(
            "=" * 70
        )


        return render_template(
            "consult_success.html",
            filename=None,
            error=str(e)
        )