import streamlit as st
import asyncio
from gpt_researcher import GPTResearcher
from gpt_researcher.utils.enum import ReportType, Tone

import pandas as pd
import os
import glob
import numpy as np
import base64
import re


import json
import re
import pandas as pd
from openai import OpenAI
import matplotlib.pyplot as plt
import plotly.graph_objects as go

# -------------------------------------------------------------------
# REPLACE HARDCODED CREDENTIALS WITH ENVIRONMENT VARIABLES
# -------------------------------------------------------------------
# Make sure you have: OPENAI_API_KEY set in your environment
API_KEY = os.getenv("OPENR", "")

site_url = "<YOUR_SITE_URL>"
site_name = "<YOUR_SITE_NAME>"

SITE_URL = site_url
SITE_NAME = site_name

output_base_name = "detailed_risk_report"  # Common base name for combined reports

def extract_top_10_risks(research_text: str, api_key: str, site_url: str, site_name: str) -> list:
    """
    Step 1: Extract the top 10 risks from the provided research text.
    The API is prompted to output a JSON object with a key "risks" containing a list of risk names.
    """
    prompt = f"""
Extract the top 10 risks mentioned in the following research content.
Provide your answer as a JSON object with a key "risks" that contains a list of strings.
Research content:
\"\"\"{research_text}\"\"\"
"""
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )
    response = client.chat.completions.create(
        extra_headers={
            "HTTP-Referer": site_url,
            "X-Title": site_name,
        },
        model="deepseek/deepseek-chat",
        messages=[{"role": "user", "content": prompt}]
    )
    result = response.choices[0].message.content
    print("Extract Top 10 Risks Response:")
    print(result)
    
    # Attempt to extract JSON from the result.
    try:
        data = json.loads(result)
    except Exception as e:
        json_match = re.search(r'({.*})', result, re.DOTALL)
        if json_match:
            try:
                data = json.loads(json_match.group(1))
            except Exception as e2:
                print("Failed to parse JSON:", e2)
                return []
        else:
            print("No JSON object found.")
            return []
    
    risks = data.get("risks", [])
    if not risks:
        print("No risks found in extracted JSON.")
    return risks

def generate_detailed_risk_report(
    risks: list,
    research_text: str,
    api_key: str,
    site_url: str,
    site_name: str,
    output_base_name: str = "detailed_risk_report"
) -> None:
    """
    Step 2: Generate a detailed risk report based on the extracted risks and research text.
    For each risk in the provided list, the API is prompted to provide:
      a) A detailed written justification explaining its significance.
      b) An exposure rating (1 to 10) with an explanation.
      c) A brief analysis and recommended internal audit procedures.
    
    The output includes two parts:
      1. A combined written document that explains in detail each risk,
         including the analysis and internal audit procedure recommendations.
      2. A JSON object with key "risks" containing a list of risk items. 
         Each risk item must have the keys:
            - "Risk"
            - "Exposure_Rating"
            - "Justification"
    
    The research text is provided for context.
    The combined report is saved to:
      - {output_base_name}.json (JSON format)
      - {output_base_name}.txt (full combined text report)
    """
    risks_str = ", ".join(risks)
    prompt = f"""
Using the following research content and the extracted risks list, generate a detailed risk report from the point of view of Internal Audit.
The sectors are Energy Distribution and Universal Banking in Central Europe.

For each risk in the following list: [{risks_str}], provide:
    a) A detailed written justification explaining the significance of the risk.
    b) An exposure rating on a scale from 1 (lowest exposure) to 10 (highest exposure), including an explanation.
    c) A brief analysis of what the risk comprises and recommended internal audit procedures to assess and mitigate the risk.

Your output should include two parts:
1. A combined written document that explains in detail each risk, including the analysis and internal audit procedure recommendations.
2. A JSON object that contains a key "risks" which is a list of risk items. Each risk item must have the keys:
    - "Risk"
    - "Exposure_Rating"
    - "Justification"

Use the following research content for reference:
\"\"\"{research_text}\"\"\"
"""
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )
    response = client.chat.completions.create(
        extra_headers={
            "HTTP-Referer": site_url,
            "X-Title": site_name,
        },
        model="deepseek/deepseek-chat",
        messages=[{"role": "user", "content": prompt}]
    )
    result = response.choices[0].message.content
    print("Detailed Risk Report Response:")
    print(result)
    
    # Extract JSON portion from the result.
    try:
        data = json.loads(result)
    except Exception as e:
        json_match = re.search(r'({.*})', result, re.DOTALL)
        if json_match:
            try:
                data = json.loads(json_match.group(1))
            except Exception as e2:
                print("Failed to parse JSON in detailed report:", e2)
                print("Response content:", result)
                return
        else:
            print("No JSON object found in detailed report.")
            print("Response content:", result)
            return
    
    json_file = output_base_name + ".json"
    txt_file = output_base_name + ".txt"
    
    # Save the detailed JSON risk report.
    with open(json_file, "w") as f:
        json.dump(data, f, indent=4)
    print(f"Detailed JSON risk report saved to {json_file}")
    
    # Also save the full combined text report for future reference.
    with open(txt_file, "w") as f:
        f.write(result)
    print(f"Combined detailed risk report saved to {txt_file}")

def generate_risk_excel_report(
    input_json_file: str = "detailed_risk_report.json",
    output_excel_file: str = "detailed_risk_report.xlsx"
) -> None:
    """
    Loads the JSON detailed risk report and processes it to:
      - Reassign each risk a unique Exposure_Rating from 1 (lowest) to 10 (highest).
      - Convert the list of risks into a pandas DataFrame.
      - Export the DataFrame to an Excel file.
    
    The expected JSON format is a key "risks" containing a list of risk items with:
      - "Risk"
      - "Exposure_Rating"
      - "Justification"
    """
    try:
        with open(input_json_file, "r") as f:
            data = json.load(f)
    except Exception as e:
        print("Error loading JSON file:", e)
        return
    
    risks_list = data.get("risks", [])
    if not risks_list:
        print("No risks found in the JSON data.")
        return

    # Helper to safely extract numeric ratings.
    def safe_rating(risk_item):
        try:
            return int(risk_item.get("Exposure_Rating", 0))
        except:
            return 0

    # Sort and reassign unique ratings from 1 to 10.
    risks_list_sorted = sorted(risks_list, key=lambda x: safe_rating(x))
    for i, risk_item in enumerate(risks_list_sorted):
        risk_item["Exposure_Rating"] = i + 1

    df = pd.DataFrame(risks_list_sorted)
    df.to_excel(output_excel_file, index=False)
    print(f"Excel table saved to {output_excel_file}")
    print("Generated Risk Report DataFrame:")
    print(df)

def get_detailed_justification(
    risk: str,
    current_justification: str,
    research_text: str,
    api_key: str,
    site_url: str,
    site_name: str
) -> str:
    """
    Uses the research text and the current justification as context to obtain a more detailed justification
    for the given risk from the perspective of Internal Audit.
    """
    prompt = f"""
Using the following research content:
\"\"\"{research_text}\"\"\"
and the current justification:
\"\"\"{current_justification}\"\"\"
please provide a more detailed and comprehensive justification for the risk: "{risk}".
Include additional analysis, context, and recommended internal audit procedures.
Your answer should be very detailed.
"""
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )
    response = client.chat.completions.create(
        extra_headers={
            "HTTP-Referer": site_url,
            "X-Title": site_name,
        },
        # You can adjust the model as needed:
        model="liquid/lfm-7b",
        messages=[{"role": "user", "content": prompt}]
    )
    result = response.choices[0].message.content
    return result

def generate_individual_reports(
    input_json_file: str,
    research_text: str,
    api_key: str,
    site_url: str,
    site_name: str,
    output_folder: str = "individual_risk_reports"
) -> None:
    """
    Generates individual text reports for each risk based on the detailed risk JSON report.
    For each risk, a more detailed justification is generated by utilizing the research context.
    Each file will include:
      - The risk name.
      - The exposure rating.
      - An enhanced detailed justification.
    The reports are saved into the specified output folder.
    """
    # Ensure the output folder exists.
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        with open(input_json_file, "r") as f:
            data = json.load(f)
    except Exception as e:
        print("Error loading JSON file for individual reports:", e)
        return
    
    risks_list = data.get("risks", [])
    if not risks_list:
        print("No risks found in the JSON data for individual reports.")
        return
    
    for risk in risks_list:
        risk_name = risk.get("Risk", "Unknown_Risk").replace(" ", "_")
        exposure = risk.get("Exposure_Rating", "N/A")
        current_justification = risk.get("Justification", "No justification provided.")
        
        # Get a more detailed justification using research_text and the current justification.
        enhanced_justification = get_detailed_justification(
            risk.get("Risk", "Unknown Risk"),
            current_justification,
            research_text,
            api_key,
            site_url,
            site_name
        )
        
        # Build the individual report text.
        report_text = f"Risk: {risk.get('Risk', 'Unknown Risk')}\n"
        report_text += f"Exposure Rating: {exposure}\n\n"
        report_text += "Enhanced Detailed Justification:\n"
        report_text += f"{enhanced_justification}\n"
        
        # Save each risk report into a separate text file.
        file_name = f"{output_folder}/risk_report_{risk_name}.txt"
        with open(file_name, "w", encoding="utf-8") as f:
            f.write(report_text)
        print(f"Individual report saved to {file_name}")


###############################################################################
# MAIN STREAMLIT APP
###############################################################################

OUTPUT_BASE_NAME = "detailed_risk_report"  # Base name for generated reports
EXCEL_FILE = OUTPUT_BASE_NAME + ".xlsx"

def run_research(query: str):
    """
    Wrapper for GPTResearcher usage. Adjust or replace as needed to
    capture the 'research_text' from GPTResearcher.
    """
    researcher = GPTResearcher(
        query=query,
        report_type="summary",  # Triggers deep research mode
    )
    research_data = asyncio.run(researcher.conduct_research())
    # If research_data is a dict, try to get "research_text", else treat it as a string
    if isinstance(research_data, dict):
        research_text = research_data.get("research_text", "")
    else:
        research_text = research_data
    return research_text, research_data

def plot_multi_series_radar(labels, data_dict, title="Radar Chart", max_val=10):
    """
    Generates a static radar chart (2D image) using Matplotlib.
    """
    N = len(labels)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    
    for metric_name, values in data_dict.items():
        values_circular = values + values[:1]
        ax.plot(angles, values_circular, linewidth=2, label=metric_name)
        ax.fill(angles, values_circular, alpha=0.1)
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    step = max_val // 5 if max_val >= 5 else 1
    radial_ticks = np.arange(0, max_val + step, step)
    ax.set_rgrids(radial_ticks, angle=0)
    ax.set_ylim(0, max_val)
    ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
    ax.set_title(title, y=1.08, fontsize=14)
    
    plt.tight_layout()
    return fig

def plot_interactive_multi_series_radar(labels, data_dict, title="Radar Chart", max_val=10, show_sliders=True):
    """
    Generates an interactive radar chart using Plotly with optional sliders.
    """
    categories = labels + [labels[0]]
    fig = go.Figure()

    for metric_name, values in data_dict.items():
        # Align lengths
        if len(values) < len(labels):
            st.warning(f"Metric '{metric_name}' has fewer values than labels. Padding with None.")
            values.extend([None] * (len(labels) - len(values)))
        elif len(values) > len(labels):
            st.warning(f"Metric '{metric_name}' has more values than labels. Truncating.")
            values = values[:len(labels)]

        values_circular = values + [values[0]]
        fig.add_trace(go.Scatterpolar(
            r=values_circular,
            theta=categories,
            fill='toself',
            name=metric_name,
            line=dict(width=2)
        ))
    
    initial_rotation = 0
    initial_radial_range = [0, max_val]
    layout_update = dict(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=initial_radial_range,
                tickfont=dict(size=12),
                gridcolor='lightgrey',
                gridwidth=1
            ),
            angularaxis=dict(
                tickfont=dict(size=12),
                gridcolor='lightgrey',
                rotation=initial_rotation
            )
        ),
        title=dict(text=title, x=0.5, font=dict(size=16)),
        showlegend=True,
        legend=dict(font=dict(size=12)),
        margin=dict(l=60, r=60, t=80, b=80),
        template="plotly_white"
    )

    if show_sliders:
        sliders = []
        rotation_steps = []
        for angle in range(0, 360, 15):
            step = dict(
                method="relayout",
                args=[{"polar.angularaxis.rotation": angle}],
                label=str(angle) + "°"
            )
            rotation_steps.append(step)
        sliders.append(dict(
            active=initial_rotation // 15,
            currentvalue={"prefix": "Rotation: ", "suffix": "°", "font": {"size": 14}},
            pad={"t": 50, "b": 10},
            x=0.1,
            len=0.8,
            y=-0.1,
            yanchor="top",
            xanchor="left",
            steps=rotation_steps
        ))

        zoom_steps = []
        min_zoom_range = max(1, max_val / 4)
        max_zoom_range = max_val * 2
        num_zoom_steps = 10
        zoom_range_values = np.linspace(min_zoom_range, max_zoom_range, num_zoom_steps)
        initial_zoom_index = (np.abs(zoom_range_values - max_val)).argmin()

        for r_max in zoom_range_values:
            step = dict(
                method="relayout",
                args=[{"polar.radialaxis.range": [0, r_max]}],
                label=f"{r_max:.1f}"
            )
            zoom_steps.append(step)

        sliders.append(dict(
            active=initial_zoom_index,
            currentvalue={"prefix": "Max Value (Zoom): ", "font": {"size": 14}},
            pad={"t": 100, "b": 10},
            x=0.1,
            len=0.8,
            y=-0.2,
            yanchor="top",
            xanchor="left",
            steps=zoom_steps
        ))
        layout_update["sliders"] = sliders

    fig.update_layout(**layout_update)
    return fig

def display_pdf(file_path: str) -> None:
    """
    Helper function to display a PDF file inline using an iframe.
    """
    with open(file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode("utf-8")
    pdf_display = (
        f'<iframe src="data:application/pdf;base64,{base64_pdf}" '
        f'width="700" height="1000" type="application/pdf"></iframe>'
    )
    st.markdown(pdf_display, unsafe_allow_html=True)

def main():
    st.title("AI Droid")
    
    # Create tabs for each section.
    tabs = st.tabs([
        "Excel Report", 
        "Research & Risk Reports", 
        "Spider Charts", 
        "Risk Coverage Analysis", 
        "Optimal Audit Plan", 
        "Individual Reports"
    ])
    
    # ------------------------------
    # Tab 1: Excel Report
    # ------------------------------
    with tabs[0]:
        st.header("Edit and Update Excel Report")
        updated_file = "updated_" + EXCEL_FILE
        if os.path.exists(updated_file):
            excel_file_to_use = updated_file
        elif os.path.exists(EXCEL_FILE):
            excel_file_to_use = EXCEL_FILE
        else:
            excel_file_to_use = None

        if excel_file_to_use:
            try:
                df = pd.read_excel(excel_file_to_use)
                required_columns = [
                    "Number of audits", 
                    "Number of man-days", 
                    "Number of high-significance findings", 
                    "Number of medium-significance findings"
                ]
                # Ensure these columns exist
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = 0

                st.write("Edit the values for each risk below:")
                edited_df = st.data_editor(df, num_rows="dynamic")
                
                if st.button("Save Updated Excel Report"):
                    updated_file = "updated_" + EXCEL_FILE
                    edited_df.to_excel(updated_file, index=False)
                    st.success(f"Updated Excel report saved as: {updated_file}")
                    with open(updated_file, "rb") as f:
                        st.download_button("Download Updated Excel Report", f, file_name=updated_file)
                
                if st.button("Reset Excel Report"):
                    try:
                        current_df = pd.read_excel(excel_file_to_use)
                        empty_df = pd.DataFrame(columns=current_df.columns)
                        empty_df.to_excel(excel_file_to_use, index=False)
                        st.success("Excel report has been reset! Only the header row remains.")
                    except Exception as e:
                        st.error(f"Error resetting Excel report: {e}")
            except Exception as e:
                st.error(f"Error loading or editing Excel report: {e}")
        else:
            st.info("Excel report not found. Please generate the research report first to create one.")

    # ------------------------------
    # Tab 2: Research & Risk Reports
    # ------------------------------
    with tabs[1]:
        st.header("Generate Research and Risk Reports")
        query = st.text_input("Enter your research query", value="What are the latest developments in quantum computing?")
        if st.button("Generate Research Report"):
            st.info("Conducting research... please wait.")
            research_text, research_data = run_research(query)
            st.subheader("Generated Research Text")
            st.text_area("Research Text", research_text, height=600)
            
            top_risks = extract_top_10_risks(research_text, API_KEY, SITE_URL, SITE_NAME)
            st.subheader("Extracted Top 10 Risks")
            st.write(top_risks)
            
            generate_detailed_risk_report(
                top_risks, research_text, API_KEY, SITE_URL, SITE_NAME, OUTPUT_BASE_NAME
            )
            st.success("Detailed risk report generated!")
            
            new_excel_file = "new_" + EXCEL_FILE
            generate_risk_excel_report(
                input_json_file=OUTPUT_BASE_NAME + ".json",
                output_excel_file=new_excel_file
            )
            st.success("New Excel report generated!")
            
            existing_file = None
            if os.path.exists("updated_" + EXCEL_FILE):
                existing_file = "updated_" + EXCEL_FILE
            elif os.path.exists(EXCEL_FILE):
                existing_file = EXCEL_FILE

            # Merge the newly generated Excel data with the existing one, if any.
            if existing_file:
                try:
                    df_existing = pd.read_excel(existing_file)
                    df_new = pd.read_excel(new_excel_file)
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                    updated_excel_file = "updated_" + EXCEL_FILE
                    df_combined.to_excel(updated_excel_file, index=False)
                    st.success("Excel report updated with new info!")
                    with open(updated_excel_file, "rb") as f:
                        st.download_button("Download Updated Excel Report", f, file_name=updated_excel_file)
                except Exception as e:
                    st.error(f"Error merging Excel reports: {e}")
            else:
                # If no existing file, rename new_... to the main excel filename
                os.rename(new_excel_file, EXCEL_FILE)
                st.success("Excel report created!")
                with open(EXCEL_FILE, "rb") as f:
                    st.download_button("Download Excel Report", f, file_name=EXCEL_FILE)
            
            # Generate individual risk reports.
            generate_individual_reports(
                input_json_file=OUTPUT_BASE_NAME + ".json",
                research_text=research_text,
                api_key=API_KEY,
                site_url=SITE_URL,
                site_name=SITE_NAME
            )
            st.success("Individual risk reports generated!")

    # ------------------------------
    # Tab 3: Spider Charts
    # ------------------------------
    with tabs[2]:
        st.header("Spider (Radar) Charts")
        updated_file = "updated_" + EXCEL_FILE
        if os.path.exists(updated_file):
            excel_file_to_use = updated_file
        elif os.path.exists(EXCEL_FILE):
            excel_file_to_use = EXCEL_FILE
        else:
            excel_file_to_use = None

        if excel_file_to_use:
            try:
                df_updated = pd.read_excel(excel_file_to_use)
                
                # Compute or rename columns needed for radar charts
                if "External Risk Exposure Index" in df_updated.columns:
                    df_updated["External_Risk_Exposure"] = df_updated["External Risk Exposure Index"]
                else:
                    max_audits = df_updated["Number of audits"].max()
                    df_updated["External_Risk_Exposure"] = (
                        df_updated["Number of audits"] / max_audits if max_audits > 0 else 0
                    )

                if "Internal Risk Propensity" in df_updated.columns:
                    df_updated["Internal_Risk_Propensity"] = df_updated["Internal Risk Propensity"]
                else:
                    total_findings = (
                        df_updated["Number of high-significance findings"] 
                        + df_updated["Number of medium-significance findings"]
                    )
                    max_findings = total_findings.max()
                    df_updated["Internal_Risk_Propensity"] = (
                        total_findings / max_findings if max_findings > 0 else 0
                    )

                if "Overall Compound Risk Index" in df_updated.columns:
                    df_updated["Overall_Compound_Risk"] = df_updated["Overall Compound Risk Index"]
                else:
                    df_updated["Overall_Compound_Risk"] = (
                        df_updated["External_Risk_Exposure"] + df_updated["Internal_Risk_Propensity"]
                    ) / 2

                if "Normalized Man-Days" in df_updated.columns:
                    df_updated["Normalized_Man_Days"] = df_updated["Normalized Man-Days"]
                else:
                    max_man_days = df_updated["Number of man-days"].max()
                    df_updated["Normalized_Man_Days"] = (
                        df_updated["Number of man-days"] / max_man_days if max_man_days > 0 else 0
                    )

                labels = [str(i + 1) for i in range(len(df_updated))]
                external_values = [v * 10 for v in df_updated["External_Risk_Exposure"].tolist()]
                internal_values = [v * 10 for v in df_updated["Internal_Risk_Propensity"].tolist()]
                compound_values = [v * 10 for v in df_updated["Overall_Compound_Risk"].tolist()]
                man_days_values = [v * 10 for v in df_updated["Normalized_Man_Days"].tolist()]

                chart_type = st.radio("Select Chart Type", ["Interactive", "Static (2D Image)"])
                if chart_type == "Interactive":
                    data_dict_chart1 = {
                        "External Risk Exposure": external_values,
                        "Internal Risk Propensity": internal_values,
                        "Overall Compound Risk": compound_values
                    }
                    fig1 = plot_interactive_multi_series_radar(
                        labels=labels,
                        data_dict=data_dict_chart1,
                        title="Risk Index Overview",
                        max_val=10
                    )
                    st.plotly_chart(fig1, use_container_width=True)

                    data_dict_chart2 = {
                        "Overall Compound Risk": compound_values,
                        "Normalized Man-Days": man_days_values
                    }
                    fig2 = plot_interactive_multi_series_radar(
                        labels=labels,
                        data_dict=data_dict_chart2,
                        title="Compound Risk vs. Man-Days",
                        max_val=10
                    )
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    # Static radar charts
                    data_dict_chart1 = {
                        "External Risk Exposure": external_values,
                        "Internal Risk Propensity": internal_values,
                        "Overall Compound Risk": compound_values
                    }
                    fig_static1 = plot_multi_series_radar(
                        labels, data_dict_chart1, title="Risk Index Overview", max_val=10
                    )
                    st.pyplot(fig_static1)

                    data_dict_chart2 = {
                        "Overall Compound Risk": compound_values,
                        "Normalized Man-Days": man_days_values
                    }
                    fig_static2 = plot_multi_series_radar(
                        labels, data_dict_chart2, title="Compound Risk vs. Man-Days", max_val=10
                    )
                    st.pyplot(fig_static2)
            except Exception as e:
                st.error(f"Error generating spider charts: {e}")
        else:
            st.info("Excel report not found. Please generate and/or update the Excel report first.")

    # ------------------------------
    # Tab 4: Risk Coverage Analysis
    # ------------------------------
    with tabs[3]:
        st.header("Risk Coverage Analysis")
        if excel_file_to_use:
            try:
                df_updated = pd.read_excel(excel_file_to_use)
                req_cols = [
                    "Number of audits", 
                    "Number of man-days", 
                    "Number of high-significance findings", 
                    "Number of medium-significance findings"
                ]
                for col in req_cols:
                    if col not in df_updated.columns:
                        st.error(f"Missing required column: {col}")
                        st.stop()

                if "External Risk Exposure Index" in df_updated.columns:
                    df_updated["External_Risk_Level"] = df_updated["External Risk Exposure Index"]
                else:
                    max_audits = df_updated["Number of audits"].max()
                    df_updated["External_Risk_Level"] = (
                        df_updated["Number of audits"] / max_audits if max_audits != 0 else 0
                    )

                df_updated["Total_Findings"] = (
                    df_updated["Number of high-significance findings"] 
                    + df_updated["Number of medium-significance findings"]
                )
                df_updated["Raw_Internal_Risk_Level"] = (
                    df_updated["Total_Findings"] / df_updated["Number of man-days"]
                )
                max_internal = df_updated["Raw_Internal_Risk_Level"].max()
                df_updated["Internal_Risk_Level"] = (
                    df_updated["Raw_Internal_Risk_Level"] / max_internal if max_internal != 0 else 0
                )
                
                df_updated["Coverage_Ratio"] = df_updated.apply(
                    lambda row: row["Internal_Risk_Level"] / row["External_Risk_Level"] 
                    if row["External_Risk_Level"] != 0 else float('inf'),
                    axis=1
                )
                
                def classify_coverage(ratio):
                    if ratio < 0.8:
                        return "Insufficiently Audited"
                    elif ratio > 1.2:
                        return "Excessively Audited"
                    else:
                        return "Sufficiently Audited"
                
                df_updated["Audit_Coverage"] = df_updated["Coverage_Ratio"].apply(classify_coverage)
                
                st.subheader("Risk Coverage Analysis Results")
                st.write(df_updated[[
                    "External_Risk_Level", "Internal_Risk_Level", 
                    "Coverage_Ratio", "Audit_Coverage"
                ]])
            except Exception as e:
                st.error(f"Error in risk coverage analysis: {e}")
        else:
            st.info("Excel report not found. Please generate and/or update the Excel report first.")

    # ------------------------------
    # Tab 5: Optimal Internal Audit Plan
    # ------------------------------
    with tabs[4]:
        st.header("Optimal Internal Audit Plan for Next 3 Years")
        if excel_file_to_use:
            try:
                df_updated = pd.read_excel(excel_file_to_use)
                # Check or compute needed columns
                if "External_Risk_Exposure" not in df_updated.columns:
                    if "External Risk Exposure Index" in df_updated.columns:
                        df_updated["External_Risk_Exposure"] = df_updated["External Risk Exposure Index"]
                    else:   
                        max_audits = df_updated["Number of audits"].max()
                        df_updated["External_Risk_Exposure"] = (
                            df_updated["Number of audits"] / max_audits if max_audits != 0 else 0
                        )

                if "Internal_Risk_Propensity" not in df_updated.columns:
                    if "Internal Risk Propensity" in df_updated.columns:
                        df_updated["Internal_Risk_Propensity"] = df_updated["Internal Risk Propensity"]
                    else:
                        total_findings = (
                            df_updated["Number of high-significance findings"] 
                            + df_updated["Number of medium-significance findings"]
                        )
                        max_findings = total_findings.max()
                        df_updated["Internal_Risk_Propensity"] = (
                            total_findings / max_findings if max_findings != 0 else 0
                        )

                # Example total capacity for the next 3 years
                CURRENT_CAPACITY = 1000

                if "Overall Compound Risk Index" in df_updated.columns:
                    df_updated["Overall_Compound_Risk"] = df_updated["Overall Compound Risk Index"]
                else:
                    df_updated["Overall_Compound_Risk"] = (
                        df_updated["External_Risk_Exposure"] + df_updated["Internal_Risk_Propensity"]
                    ) / 2

                df_updated["Overall_Compound_Risk"] = pd.to_numeric(
                    df_updated["Overall_Compound_Risk"], errors='coerce'
                ).fillna(0)

                total_risk = df_updated["Overall_Compound_Risk"].sum()
                if total_risk == 0:
                    st.error("Total risk is zero; cannot compute optimal allocation.")
                else:
                    df_updated["Risk_Score"] = (
                        df_updated["Overall_Compound_Risk"] / total_risk
                    ).replace([np.inf, -np.inf], 0).fillna(0)

                    # Distribute audits based on Risk_Score
                    df_updated["Optimal_Audits"] = (df_updated["Risk_Score"] * CURRENT_CAPACITY).round().astype(int)

                    # Ensure every risk with a nonzero risk level has at least 1 audit
                    df_updated.loc[df_updated["Overall_Compound_Risk"] > 0, "Optimal_Audits"] = \
                        df_updated.loc[df_updated["Overall_Compound_Risk"] > 0, "Optimal_Audits"].apply(lambda x: max(x, 1))
                    
                    allocated = df_updated["Optimal_Audits"].sum()
                    difference = CURRENT_CAPACITY - allocated
                    # If we haven't allocated exactly the total capacity, fix by adjusting 
                    # the row with the largest risk
                    if difference != 0:
                        max_index = df_updated["Overall_Compound_Risk"].idxmax()
                        df_updated.at[max_index, "Optimal_Audits"] += difference

                    optimal_plan_file = "optimal_internal_audit_plan.xlsx"
                    df_updated.to_excel(optimal_plan_file, index=False)
                    
                    st.subheader("Optimal Internal Audit Plan Generated")
                    st.write(df_updated[["Optimal_Audits", "Risk_Score", "Overall_Compound_Risk"]])
                    with open(optimal_plan_file, "rb") as f:
                        st.download_button("Download Optimal Internal Audit Plan", f, file_name=optimal_plan_file)
            except Exception as e:
                st.error(f"Error generating optimal internal audit plan: {e}")
        else:
            st.info("Excel report not found. Please generate and/or update the Excel report first.")

    # ------------------------------
    # Tab 6: Individual Reports
    # ------------------------------
    with tabs[5]:
        st.header("Individual Risk Reports")

        # Locate generated text reports
        risk_report_files = glob.glob(os.path.join("individual_risk_reports", "**", "*.txt"), recursive=True)
        st.write("Found risk report files:", risk_report_files)
        
        if risk_report_files:
            # Create tab labels as numbers (as strings)
            tab_labels = [str(i + 1) for i in range(len(risk_report_files))]
            sub_tabs = st.tabs(tab_labels)
            
            # Display each risk report in its corresponding tab.
            for i, sub_tab in enumerate(sub_tabs):
                with sub_tab:
                    try:
                        with open(risk_report_files[i], "r") as report_file:
                            report_content = report_file.read()
                        st.text_area(f"Risk Report {i + 1}", report_content, height=200)
                    except Exception as e:
                        st.write(f"Error loading {risk_report_files[i]}: {e}")
        else:
            st.info("No individual risk reports found. Please generate them first.")

        st.markdown("---")
        st.info("Below you can download and view individual risk reports as PDF/TXT files, if any are generated.")
        
        # If the user generated PDF versions separately, you'd see them here:
        reports_dir = "individual_risk_reports"
        if os.path.isdir(reports_dir):
            pdf_files = glob.glob(os.path.join(reports_dir, "*.pdf"))
            txt_files = glob.glob(os.path.join(reports_dir, "*.txt"))
            
            if pdf_files or txt_files:
                if pdf_files:
                    st.subheader("PDF Reports")
                    # Display each PDF with a download button and a view option.
                    for file in pdf_files:
                        col1, col2 = st.columns([1, 1])
                        with col1:
                            with open(file, "rb") as f:
                                st.download_button(
                                    label=f"Download {os.path.basename(file)}",
                                    data=f.read(),
                                    file_name=os.path.basename(file)
                                )
                        with col2:
                            if st.button(f"View {os.path.basename(file)}", key=file):
                                display_pdf(file)
                if txt_files:
                    st.subheader("TXT Reports")
                    for file in txt_files:
                        with open(file, "rb") as f:
                            st.download_button(
                                label=f"Download {os.path.basename(file)}",
                                data=f.read(),
                                file_name=os.path.basename(file)
                            )
            else:
                st.warning("No PDF or TXT files found in the directory.")
        else:
            st.warning("The 'individual_risk_reports' folder does not exist. Ensure individual reports are generated.")

if __name__ == "__main__":
    main()
