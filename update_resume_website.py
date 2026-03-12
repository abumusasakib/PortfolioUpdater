import openpyxl
from datetime import datetime
from bs4 import BeautifulSoup
import os
import sys

# Import extraction logic from the main script
from extract_excel_data import extract_excel_data

def inspect_and_verify(excel_path):
    """Combines inspection and verification logic"""
    print(f"--- FOCUSED INSPECTION & VERIFICATION ---")
    print(f"Excel File: {excel_path}")
    
    # Inspection: Row-level check
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    target_sheets = ['01-personal', '02-career-highlight', '06-job-history', '12-contact']
    
    print("\n[Row-Level Inspection Summary]")
    for sheet_name in target_sheets:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(min_row=1, max_row=5, values_only=True))
            print(f"- Sheet '{sheet_name}': Found {ws.max_row} rows. (First 5 rows checked)")
        else:
            print(f"- Sheet '{sheet_name}': NOT FOUND!")

    # Verification: Data-level check
    data = extract_excel_data(excel_path)
    print("\n[Data-Level Verification]")
    print(f"Name extracted: {data['personal'].get('Name')}")
    print(f"Designation: {data['personal'].get('Designation')}")
    print(f"Location: {data['personal'].get('Address')}")
    print(f"Jobs found: {len(data['jobs'])}")
    print(f"Skills found: {len(data['technical_skills'])}")
    print(f"Career highlights: {len(data['career_highlights'])}")
    
    return data

def sync_common_elements(soup, data):
    """Synchronizes navbar and footer across all pages"""
    # 1. Update Footer Socials
    footer_socials = soup.select('.footer-socials a')
    linkedin = data['contact'].get('LinkedIn Page')
    github = data['contact'].get('GitHub/GitLab Page')
    facebook = data['contact'].get('Facebook Page')
    instagram = data['contact'].get('Instagram/Twitter Page')

    for a in footer_socials:
        href = a.get('href', '')
        icon = a.find('i')
        if not icon: continue
        
        cls = " ".join(icon.get('class', []))
        if 'linkedin' in cls and linkedin:
            a['href'] = linkedin
        elif 'github' in cls and github:
            a['href'] = github
        elif 'facebook' in cls and facebook:
            a['href'] = facebook
        elif 'instagram' in cls and instagram:
            a['href'] = instagram

    # 2. Update Footer Copyright Name
    footer_text = soup.select_one('.footer-text')
    if footer_text:
        year = datetime.now().year
        name = data['personal'].get('Name', 'Abu Musa Sakib')
        footer_text.string = f"© {year} {name}. All rights reserved."

def update_website(data, base_dir):
    """Updates the portfolio-sakib website files with extracted data"""
    if not os.path.exists(base_dir):
        print(f"Error: Directory {base_dir} not found!")
        return

    print(f"\n--- UPDATING WEBSITE IN {base_dir} ---")

    pages = ['index.html', 'about-me.html', 'skills.html', 'portfolio.html', 'contact.html']

    for page_name in pages:
        page_path = os.path.join(base_dir, page_name)
        if not os.path.exists(page_path):
            continue

        print(f"Processing {page_name}...")
        with open(page_path, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f, 'html.parser')

        # Sync common elements (footer, socials)
        sync_common_elements(soup, data)

        # Page-specific updates
        if page_name == 'index.html':
            name = data['personal'].get('Name', 'Abu Musa Sakib')
            designation = data['personal'].get('Designation', 'Software Developer')
            
            hero_name = soup.select_one('.heading-text .accent-text')
            if hero_name: hero_name.string = name
            
            card_title = soup.select_one('.home-card .card-title')
            if card_title: card_title.string = designation
            
            card_text = soup.select_one('.home-card .card-text')
            if card_text:
                # Fix nested <p> issue: if card_text is a <p>, turn it into a <div>
                if card_text.name == 'p':
                    card_text.name = 'div'
                card_text.clear()
                for h in data['career_highlights'][:3]:
                    p = soup.new_tag('p', attrs={'class': 'mb-2'})
                    p.string = h
                    card_text.append(p)

        elif page_name == 'about-me.html':
            # Update Education
            for card in soup.find_all(class_='card'):
                title = card.find(class_='card-title')
                if title and 'Education' in title.text:
                    body = card.find(class_='card-body')
                    hr = body.find('hr')
                    if hr:
                        # Clear old education content after hr
                        for sibling in hr.find_all_next():
                            sibling.decompose()
                        
                        container = soup.new_tag('div', attrs={'class': 'card-text'})
                        for edu in data['education']:
                            strong = soup.new_tag('strong', attrs={'class': 'd-block mb-1'})
                            strong.string = edu.get('institute', '')
                            container.append(strong)
                            
                            span = soup.new_tag('span', attrs={'class': 'text-white'})
                            span.string = str(edu.get('year', '')).split('.')[0]
                            container.append(span)
                            
                            h6 = soup.new_tag('h6')
                            h6.string = edu.get('degree', '')
                            container.append(h6)
                            
                            p = soup.new_tag('p', attrs={'class': 'mb-2'})
                            p.string = f"Subject/Discipline: {edu.get('subject', '')}"
                            container.append(p)
                            container.append(soup.new_tag('br'))
                        
                        hr.insert_after(container)
                    break
            
            # Update Personal Details
            for card in soup.find_all(class_='card'):
                title = card.find(class_='card-title')
                if title and 'Personal Details' in title.text:
                    p_text = card.select_one('.card-text')
                    if p_text:
                        addr = data['personal'].get('Address', 'N/A')
                        nat = data['personal'].get('Nationality', 'Bangladeshi')
                        p_text.clear()
                        p_text.append(f"Location: {addr}")
                        p_text.append(soup.new_tag('br'))
                        p_text.append(f"Nationality: {nat}")

        elif page_name == 'skills.html':
            thematic_images = {
                'coding': 'img/Showcase-your-Programming-Skills.jpg',
                'programming': 'img/Showcase-your-Programming-Skills.jpg',
                'vcs': 'img/Showcase-your-Programming-Skills.jpg',
                'system': 'img/LCOM_Website_DigitalLiteracy_Image1_Dec16.png',
                'communication': 'img/Lessonly_Resources_communication-skills-training_Mock-Up_F.png',
                'general': 'img/skills-word-cloud-business-concept-58626753.jpg'
            }
            skills_row = soup.select_one('.skills-row') or soup.find(class_='row')
            if skills_row:
                skills_row.clear()
                for s_entry in data['technical_skills']:
                    cat = s_entry.get('category', 'Skill')
                    cat_lower = cat.lower()
                    img_src = 'img/skills-word-cloud-business-concept-58626753.jpg'
                    for kw, src in thematic_images.items():
                        if kw in cat_lower: img_src = src; break
                    
                    card = soup.new_tag('div', attrs={'class': 'skill-card'})
                    img_cont = soup.new_tag('div', attrs={'class': 'skill-img-container'})
                    img = soup.new_tag('img', attrs={'src': img_src, 'class': 'skill-img', 'alt': cat})
                    img_cont.append(img)
                    
                    content = soup.new_tag('div', attrs={'class': 'skill-content'})
                    h3 = soup.new_tag('h3', attrs={'class': 'skill-title'})
                    h3.string = cat
                    ul = soup.new_tag('ul', attrs={'class': 'skill-list'})
                    for s in s_entry.get('skills', '').split(','):
                        if s.strip():
                            li = soup.new_tag('li')
                            li.string = s.strip()
                            ul.append(li)
                    content.append(h3)
                    content.append(ul)
                    
                    card.append(img_cont)
                    card.append(content)
                    skills_row.append(card)

        elif page_name == 'portfolio.html':
            proj_row = soup.select_one('.projects-row')
            if proj_row and data.get('projects'):
                proj_row.clear()
                for proj in data['projects']:
                    card = soup.new_tag('div', attrs={'class': 'project-card'})
                    img_cont = soup.new_tag('div', attrs={'class': 'project-img-container'})
                    img = soup.new_tag('img', attrs={'src': 'img/92220679-corporate-identity-template-on-white-paper-background-photo-of-blank-stationery-set-mockup-for-desig.jpg', 'class': 'project-img', 'alt': proj['name']})
                    img_cont.append(img)
                    
                    content = soup.new_tag('div', attrs={'class': 'project-content'})
                    h3 = soup.new_tag('h3', attrs={'class': 'project-title'})
                    h3.string = proj['name']
                    p = soup.new_tag('p', attrs={'class': 'project-description'})
                    p.string = proj['description'].strip()
                    
                    tech_div = soup.new_tag('div', attrs={'class': 'project-tech'})
                    for t in proj['tech']:
                        span = soup.new_tag('span', attrs={'class': 'tech-badge'})
                        span.string = t
                        tech_div.append(span)
                    
                    content.append(h3)
                    content.append(p)
                    content.append(tech_div)
                    
                    card.append(img_cont)
                    card.append(content)
                    proj_row.append(card)

        elif page_name == 'contact.html':
            email = data['contact'].get('Email', '')
            phone = data['contact'].get('Contact Number', '')
            if phone.startswith('='): phone = phone[1:]
            
            email_a = soup.select_one('a[href^="mailto:"]')
            if email_a:
                email_a['href'] = f"mailto:{email}"
                email_a.string = email
            
            phone_a = soup.select_one('a[href^="tel:"]')
            if phone_a:
                phone_a['href'] = f"tel:{phone.replace(' ', '')}"
                phone_a.string = phone
            
            for a in soup.select('.details-card .contact-link'):
                href = a.get('href', '').lower()
                if 'linkedin' in href: a['href'] = data['contact'].get('LinkedIn Page', '')
                elif 'github' in href: a['href'] = data['contact'].get('GitHub/GitLab Page', '')

        # Write each updated file
        with open(page_path, 'w', encoding='utf-8') as f:
            f.write(soup.prettify())

    print(f"\nWebsite updated successfully!")

if __name__ == "__main__":
    excel_file = "Résumé__Abu.Musa.Sakib.xlsx"
    portfolio_dir = "portfolio-sakib"
    
    data = inspect_and_verify(excel_file)
    update_website(data, portfolio_dir)

