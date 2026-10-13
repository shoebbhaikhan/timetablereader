import streamlit as st
import pandas as pd
import openpyxl
import datetime
import os
import re
import json

EXCLUSION_FILE = "excluded_faculty.json"

st.set_page_config(
    page_title="UID ID Timetable & Generator",
    page_icon="📅",
    layout="wide"
)

st.markdown("""
<style>
    .badge-free {
        background-color: #dcfce7;
        color: #15803d;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
        margin: 2px 0;
    }
    .badge-busy {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
        margin: 2px 0;
    }
    .faculty-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .wt-table {
        width: 100%;
        border-collapse: collapse;
        font-family: inherit;
        margin-top: 14px;
    }
    .wt-th {
        background-color: #0f172a;
        color: #ffffff;
        padding: 10px 8px;
        text-align: center;
        font-size: 0.88rem;
        border: 1px solid #334155;
    }
    .wt-td-label {
        background-color: #f8fafc;
        color: #0f172a;
        font-weight: 700;
        padding: 8px 10px;
        border: 1px solid #cbd5e1;
        font-size: 0.85rem;
        vertical-align: middle;
        white-space: nowrap;
    }
    .wt-td {
        border: 1px solid #cbd5e1;
        padding: 6px 4px;
        vertical-align: top;
        background-color: #ffffff;
        width: 18%;
    }
    .s-pill {
        border-radius: 4px;
        padding: 3px 6px;
        margin-bottom: 4px;
        font-size: 0.74rem;
        line-height: 1.25;
        display: block;
        border-left: 3px solid transparent;
    }
    .s1 { background: #dbeafe; color: #1e40af; border-color: #2563eb; }
    .s2 { background: #ccfbf1; color: #0f766e; border-color: #0d9488; }
    .s3 { background: #fef3c7; color: #92400e; border-color: #d97706; }
    .s4 { background: #f3e8ff; color: #6b21a8; border-color: #9333ea; }
    .s-off { background: #f1f5f9; color: #94a3b8; font-style: italic; }

    /* Completion Status Badges */
    .status-badge {
        padding: 4px 12px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
        text-align: center;
        letter-spacing: 0.02em;
    }
    .status-completed {
        background-color: #dcfce7;
        color: #166534;
        border: 1px solid #bbf7d0;
    }
    .status-active {
        background-color: #ffedd5;
        color: #9a3412;
        border: 1px solid #fed7aa;
    }
    .status-upcoming {
        background-color: #cffafe;
        color: #155e75;
        border: 1px solid #a5f3fc;
    }
</style>
""", unsafe_allow_html=True)


def get_excluded_faculty(target_date=None):
    """
    Returns a set of faculty names excluded on target_date.
    Supports plain string entries (permanent) or dicts with 'end_date' (YYYY-MM-DD).
    """
    if os.path.exists(EXCLUSION_FILE):
        try:
            with open(EXCLUSION_FILE, "r") as f:
                data = json.load(f)
                excluded = set()
                for entry in data:
                    if isinstance(entry, str):
                        excluded.add(entry)
                    elif isinstance(entry, dict) and "name" in entry:
                        end_date_str = entry.get("end_date")
                        if end_date_str and target_date:
                            try:
                                end_dt = datetime.date.fromisoformat(end_date_str)
                                if target_date > end_dt:
                                    excluded.add(entry["name"])
                            except ValueError:
                                excluded.add(entry["name"])
                        else:
                            excluded.add(entry["name"])
                return excluded
        except Exception:
            return set()
    return set()


COURSES_DATA = [
    # 1st Sem B.Design (PSD) Batch 2026-30
    {"code": "26UDU01001", "title": "Design Essentials", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01002", "title": "Visual Representation Skills", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01103", "title": "Materials and Craftsmanship I", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01204", "title": "Colour and Form", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01305", "title": "History of Art & Design", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01406", "title": "Fundamentals of AI", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    {"code": "26UDU01607", "title": "Digital Tools", "sem": "I", "prog_code": "1019", "batch": "1st Sem B.Des (PSD) 2026-30"},
    
    # 3rd Sem B. Design Product Design Batch 2025-29
    {"code": "UC012030001", "title": "Personality Development", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203001203", "title": "Design Research", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203006211", "title": "Product Visualization", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203006212", "title": "Form, Aesthetic and Emotion", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203006207", "title": "Studio- Human Centric Design", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203006215", "title": "Design Articulation with AI", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    {"code": "31203006214", "title": "Indian Design System", "sem": "III", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
    
    # 5th Sem B. Design Product Design Batch 2024-28
    {"code": "31305006328", "title": "CAID", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001303", "title": "Conceptualization & Characterization", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001304", "title": "Speed Modelling in Clay", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001306", "title": "Space Perception", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001324", "title": "Matte-Painting", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001325", "title": "Ad Film Production", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001326", "title": "Fashion Styling", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001327", "title": "UX/UI", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001328", "title": "The Art of Delightful Design", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305001329", "title": "Branding And Identity Design", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006322", "title": "Human Factors", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006329", "title": "Portfolio Design with Voice Agents", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006324", "title": "Studio- Humanizing Technology", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006325", "title": "Experience Design", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006326", "title": "Packaging Design", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006327", "title": "Speculative design", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    {"code": "31305006329", "title": "Portfolio with AI", "sem": "V", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
    
    # 7th Sem B.Design Product Design Batch 2023-27
    {"code": "31407001400", "title": "Design Management", "sem": "VII", "prog_code": "1019", "batch": "7th Sem B.Des PD 2023-27"},
    {"code": "31407006403", "title": "Internship", "sem": "VII", "prog_code": "1019", "batch": "7th Sem B.Des PD 2023-27"},
    {"code": "31407006405", "title": "System Analysis and Design with Voice Agents", "sem": "VII", "prog_code": "1019", "batch": "7th Sem B.Des PD 2023-27"},
    
    # 1st Sem M. Design Product & Service Design Batch 2026-28
    {"code": "26UDP01317", "title": "Design Appreciation & Storytelling", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01018", "title": "Design Foundation", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01019", "title": "Form Studies", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01020", "title": "Design Studio I", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01421", "title": "Design Prototyping", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01122", "title": "Frugal Innovation", "sem": "I", "prog_code": "1025", "batch": "1st Sem M.Des PSD 2026-28"},
    {"code": "26UDP01123",
