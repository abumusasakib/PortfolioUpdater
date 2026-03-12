import openpyxl
from datetime import datetime
from extract_excel_data import extract_excel_data

def verify_data(output_file):
    excel_path = "Résumé__Abu.Musa.Sakib.xlsx"
    
    data = extract_excel_data(excel_path)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(f"Verifying extraction from: {excel_path}\n")
        
        f.write("\n--- PERSONAL INFO ---\n")
        f.write(f"Name: {data['personal'].get('Name')}\n")
        f.write(f"Designation: {data['personal'].get('Designation')}\n")
        f.write(f"Address: {data['personal'].get('Address')}\n")
        
        f.write("\n--- CAREER HIGHLIGHTS ---\n")
        f.write(f"Total highlights: {len(data['career_highlights'])}\n")
        for h in data['career_highlights'][:3]:
            f.write(f"- {h}\n")
        
        f.write("\n--- JOBS ---\n")
        f.write(f"Total jobs: {len(data['jobs'])}\n")
        for job in data['jobs']:
            f.write(f"\nPosition: {job.get('position')}\n")
            f.write(f"Organization: {job.get('organization')}\n")
            f.write(f"Dates: {job.get('from')} to {job.get('to')}\n")
            f.write(f"Bullets: {len(job.get('bullets', []))}\n")
        
        f.write("\n--- PROJECTS ---\n")
        f.write(f"Total projects: {len(data['projects'])}\n")
        for proj in data['projects']:
            f.write(f"\nProject: {proj.get('name')}\n")
            f.write(f"Description: {proj.get('description')}\n")
            f.write(f"Tech: {', '.join(proj.get('tech', []))}\n")
            f.write(f"Tasks: {len(proj.get('tasks', []))}\n")

        f.write("\n--- CONTACT ---\n")
        for k, v in data['contact'].items():
            f.write(f"{k}: {v}\n")

if __name__ == "__main__":
    verify_data("verify_output.txt")
