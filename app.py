import streamlit as st
import pandas as pd
from main import export_to_excel

from main import (
    generate_timetable,
    sections,
    days,
    periods,
    short,
    faculty_map,
    faculty_name_map
)

st.set_page_config(
    page_title="AI Timetable Generator",
    page_icon="📅",
    layout="wide"
)

st.title("📅 AI Timetable Generator")
st.write("Generate AI-based class timetables for all sections.")

if st.button("Generate Timetable"):

    timetable = generate_timetable()

    st.success("Timetable Generated Successfully!")

    for section in sections:

        st.subheader(f"Section {section}")

        table = []

        for day in days:

            row = {"Day": day}

            for i, cid in enumerate(timetable[section][day]):

                sname = short(cid)

                fid = faculty_map.get((cid, section), "")

                fname = faculty_name_map.get(fid, "")

                if fname != "":
                    row[periods[i]] = f"{sname} ({fname})"
                else:
                    row[periods[i]] = sname

            table.append(row)

        df = pd.DataFrame(table)

        st.dataframe(df, use_container_width=True)

    export_to_excel(timetable, "AI_Timetable_Output.xlsx")

    with open("AI_Timetable_Output.xlsx", "rb") as file:
        st.download_button(
            label="📥 Download Timetable Excel",
            data=file,
            file_name="AI_Timetable_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )