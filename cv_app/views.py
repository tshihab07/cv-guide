import re
from copy import deepcopy
from io import BytesIO

from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.forms import modelformset_factory
from django.http import HttpResponse
from django.shortcuts import redirect, render
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer

from .forms import (
    BenefitItemForm,
    EducationForm,
    ExperienceForm,
    FAQItemForm,
    FeatureItemForm,
    LandingPageContentForm,
    ProcessStepForm,
    ProjectForm,
    SkillCategoryForm,
    TestimonialForm,
)
from .models import (
    BenefitItem,
    FAQItem,
    FeatureItem,
    LandingPageContent,
    ProcessStep,
    Testimonial,
)

EMPTY_BUILDER_DATA = {
    'personal': {
        'full_name': '',
        'email': '',
        'phone': '',
        'linkedin': '',
        'github': '',
        'portfolio': '',
        'location': '',
    },
    'job_requirements': {
        'job_title': '',
        'role': '',
        'job_description': '',
        'keywords': '',
        'required_skills': '',
    },
    'skill_categories': [],
    'projects': [],
    'education': [],
    'experience': [],
}

DEFAULT_SKILL_CATEGORIES = [
    'Programming Languages',
    'Machine Learning',
    'Data Analysis & Visualization',
    'Database Management',
    'Backend Development',
    'Cloud & Deployment',
    'Web Scraping & Automation',
    'Dev Tools',
    'Software Testing & QA',
]

SKILL_KEYWORD_MAP = {
    'Programming Languages': [
        ('python', 'Python'),
        ('c++', 'C++'),
        ('c', 'C'),
        ('java', 'Java'),
        ('javascript', 'JavaScript'),
        (' js', 'JavaScript'),
        ('typescript', 'TypeScript'),
        ('php', 'PHP'),
        ('ruby', 'Ruby'),
        ('go', 'Go'),
        ('kotlin', 'Kotlin'),
        ('swift', 'Swift'),
        ('sql', 'SQL'),
        ('bash', 'Bash'),
        ('shell', 'Shell'),
        ('r', 'R'),
        ('scala', 'Scala'),
        ('rust', 'Rust'),
        ('matlab', 'MATLAB'),
    ],
    'Machine Learning': [
        ('machine learning', 'Machine Learning'),
        ('ml', 'Machine Learning'),
        ('deep learning', 'Deep Learning'),
        ('artificial intelligence', 'Artificial Intelligence'),
        (' ai ', 'Artificial Intelligence'),
        ('scikit-learn', 'Scikit-learn'),
        ('tensorflow', 'TensorFlow'),
        ('pytorch', 'PyTorch'),
        ('keras', 'Keras'),
        ('feature engineering', 'Feature Engineering'),
        ('data preprocessing', 'Data Preprocessing'),
        ('feature scaling', 'Feature Scaling and Normalization'),
        ('normalization', 'Feature Scaling and Normalization'),
        ('hyperparameter tuning', 'Hyperparameter Tuning'),
        ('model validation', 'Model Validation'),
        ('mlops', 'MLOps'),
        ('ci/cd', 'CI / CD'),
        ('computer vision', 'Computer Vision'),
        ('nlp', 'NLP'),
        ('natural language processing', 'NLP'),
        ('reinforcement learning', 'Reinforcement Learning'),
        ('xgboost', 'XGBoost'),
        ('lightgbm', 'LightGBM'),
        ('optuna', 'Optuna'),
        ('docker', 'Docker'),
    ],
    'Data Analysis & Visualization': [
        ('pandas', 'Pandas'),
        ('numpy', 'NumPy'),
        ('matplotlib', 'Matplotlib'),
        ('seaborn', 'Seaborn'),
        ('tableau', 'Tableau'),
        ('power bi', 'Power BI'),
        ('metabase', 'Metabase'),
        ('eda', 'EDA'),
        ('data analysis', 'Data Analysis'),
        ('analytics', 'Analytics'),
        ('visualization', 'Data Visualization'),
        ('plotly', 'Plotly'),
        ('ggplot', 'ggplot'),
        ('superset', 'Apache Superset'),
    ],
    'Database Management': [
        ('postgresql', 'PostgreSQL'),
        ('mysql', 'MySQL'),
        ('mongodb', 'MongoDB'),
        ('sqlite', 'SQLite'),
        ('nosql', 'NoSQL'),
        ('database', 'Databases'),
        ('sql', 'SQL'),
        ('redis', 'Redis'),
        ('cassandra', 'Cassandra'),
        ('dynamodb', 'DynamoDB'),
        ('oracle', 'Oracle'),
    ],
    'Backend Development': [
        ('django', 'Django'),
        ('flask', 'Flask'),
        ('fastapi', 'FastAPI'),
        ('rest api', 'REST APIs'),
        ('restful', 'REST APIs'),
        ('api', 'API Development'),
        ('backend', 'Backend Development'),
        ('authentication', 'Authentication'),
        ('oauth', 'OAuth'),
        ('jwt', 'JWT'),
        ('node.js', 'Node.js'),
        ('nodejs', 'Node.js'),
        ('express', 'Express.js'),
        ('spring boot', 'Spring Boot'),
        ('graphql', 'GraphQL'),
        ('grpc', 'gRPC'),
        ('microservices', 'Microservices'),
    ],
    'Cloud & Deployment': [
        ('aws', 'AWS'),
        ('gcp', 'Google Cloud Platform'),
        ('google cloud', 'Google Cloud Platform'),
        ('azure', 'Azure'),
        ('docker', 'Docker'),
        ('kubernetes', 'Kubernetes'),
        ('ci/cd', 'CI / CD'),
        ('deployment', 'Deployment'),
        ('render', 'Render'),
        ('vercel', 'Vercel'),
        ('supabase', 'Supabase'),
        ('mongodb atlas', 'MongoDB Atlas'),
        ('cloud', 'Cloud'),
        ('terraform', 'Terraform'),
        ('ansible', 'Ansible'),
        ('jenkins', 'Jenkins'),
        ('github actions', 'GitHub Actions'),
        ('gitlab ci', 'GitLab CI'),
    ],
    'Web Scraping & Automation': [
        ('web scraping', 'Web Scraping'),
        ('requests', 'Requests'),
        ('selenium', 'Selenium'),
        ('beautifulsoup', 'BeautifulSoup'),
        ('scrapy', 'Scrapy'),
        ('automation', 'Automation'),
        ('bots', 'Automation Bots'),
        ('playwright', 'Playwright'),
        ('puppeteer', 'Puppeteer'),
    ],
    'Dev Tools': [
        ('git', 'Git'),
        ('github', 'GitHub'),
        ('anaconda', 'Anaconda Navigator'),
        ('jupyter', 'Jupyter Notebook'),
        ('colab', 'Google Colab'),
        ('visual studio', 'Visual Studio'),
        ('vscode', 'VS Code'),
    ],
    'Software Testing & QA': [
        ('testing', 'Testing Fundamentals'),
        ('qa', 'Software Testing'),
        ('quality assurance', 'Quality Assurance'),
        ('pytest', 'Pytest'),
        ('manual testing', 'Manual Testing'),
        ('unit testing', 'Unit Testing'),
        ('bug reporting', 'Bug Reporting'),
        ('test case', 'Test Case Design'),
        ('selenium', 'Selenium Testing'),
        ('cypress', 'Cypress'),
        ('jest', 'Jest'),
    ],
    'Project Management & Soft Skills': [
        ('agile', 'Agile'),
        ('scrum', 'Scrum'),
        ('jira', 'Jira'),
        ('leadership', 'Leadership'),
        ('teamwork', 'Teamwork'),
        ('communication', 'Communication'),
        ('problem solving', 'Problem Solving'),
        ('documentation', 'Documentation'),
        ('mentoring', 'Mentoring'),
        ('kanban', 'Kanban'),
        ('confluence', 'Confluence'),
        ('trello', 'Trello'),
        ('notion', 'Notion'),
    ],
}


def _seed_default_content():
    content = LandingPageContent.objects.first()
    if not content:
        content = LandingPageContent.objects.create()

    if not FeatureItem.objects.exists():
        FeatureItem.objects.bulk_create([
            FeatureItem(title='ATS optimized', description='Generate a clean CV that is easy for recruiting systems to read.', icon='⚡', order=1),
            FeatureItem(title='Job-targeted content', description='Shape your summary and skill list around the exact role you want.', icon='🎯', order=2),
            FeatureItem(title='Instant downloads', description='Export your final CV as PDF or DOCX in one click.', icon='📄', order=3),
        ])

    if not ProcessStep.objects.exists():
        ProcessStep.objects.bulk_create([
            ProcessStep(title='Share your details', description='Enter your profile, background, and target role.', order=1),
            ProcessStep(title='Match the role', description='Add job requirements and skills to align the content.', order=2),
            ProcessStep(title='Generate and download', description='Preview the ATS-ready result and export it instantly.', order=3),
        ])

    if not BenefitItem.objects.exists():
        BenefitItem.objects.bulk_create([
            BenefitItem(title='Keyword optimization', description='Use the job description to include relevant terms naturally.', icon='🔑', order=1),
            BenefitItem(title='Professional formatting', description='Create a structured CV with strong sections and easy reading.', icon='🧩', order=2),
            BenefitItem(title='No account required', description='No signup, no login, and no long-term data storage.', icon='🔒', order=3),
        ])

    if not Testimonial.objects.exists():
        Testimonial.objects.bulk_create([
            Testimonial(quote='The output looked polished and recruiter-friendly from the very first draft.', author='Aisha K.', role='Software Engineer', order=1),
            Testimonial(quote='I used it to tailor my CV for three roles in one afternoon.', author='David J.', role='Marketing Lead', order=2),
        ])

    if not FAQItem.objects.exists():
        FAQItem.objects.bulk_create([
            FAQItem(question='Is my data safe?', answer='Yes. We do not store your information beyond the current session and you can download your CV immediately.', order=1),
            FAQItem(question='Can I export to PDF or DOCX?', answer='Yes. The preview page includes both options for quick downloads.', order=2),
        ])

    return content


def _get_builder_data(request):
    data = request.session.get('builder_data')
    if not data:
        data = deepcopy(EMPTY_BUILDER_DATA)
        request.session['builder_data'] = data
    return data


def _extract_skill_categories(job_description='', required_skills='', extra_skills='', keywords=''):
    text = ' '.join([job_description, required_skills, extra_skills, keywords]).lower()
    text = re.sub(r'[^a-z0-9+\s./-]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    categories = []
    for category in DEFAULT_SKILL_CATEGORIES:
        matches = []
        seen = set()
        for raw_term, display_term in SKILL_KEYWORD_MAP.get(category, []):
            if not raw_term:
                continue
            if raw_term in text:
                if display_term not in seen:
                    seen.add(display_term)
                    matches.append(display_term)

        if matches:
            categories.append({'category': category, 'items': ', '.join(matches)})

    if not categories:
        categories = [
            {'category': 'Project Management & Soft Skills', 'items': 'Communication, Problem Solving'},
        ]

    return categories


def _generate_summary(data):
    name = data.get('personal', {}).get('full_name') or ''
    title = data.get('job_requirements', {}).get('job_title') or data.get('job_requirements', {}).get('role') or ''
    skills = []
    for category in data.get('skill_categories', []):
        items = [item.strip() for item in category.get('items', '').split(',') if item.strip()]
        skills.extend(items)
    top_skills = ', '.join(skills[:6]) if skills else ''
    parts = []
    if title:
        parts.append(f"{title} professional with experience in {top_skills}" if top_skills else title)
    if name:
        return f"{name} — " + (parts[0] if parts else 'Experienced professional')
    return parts[0] if parts else 'Experienced professional aligned to the role.'


def _calculate_score(data):
    keywords = [item.strip().lower() for item in data['job_requirements'].get('keywords', '').split(',') if item.strip()]
    required_skills = [item.strip().lower() for item in data['job_requirements'].get('required_skills', '').split(',') if item.strip()]
    all_keywords = keywords + required_skills

    text = ' '.join([
        data['job_requirements'].get('job_title', ''),
        data['job_requirements'].get('job_description', ''),
        data['job_requirements'].get('role', ''),
        ' '.join([item.get('items', '') for item in data.get('skill_categories', [])]),
        ' '.join([project.get('title', '') + ' ' + project.get('description', '') for project in data.get('projects', [])]),
        ' '.join([edu.get('institution', '') + ' ' + edu.get('degree', '') + ' ' + edu.get('details', '') for edu in data.get('education', [])]),
        ' '.join([exp.get('company', '') + ' ' + exp.get('role', '') + ' ' + exp.get('responsibilities', '') for exp in data.get('experience', [])]),
    ]).lower()

    matched_keywords = []
    for keyword in all_keywords:
        if keyword and re.search(r'\b' + re.escape(keyword) + r'\b', text):
            matched_keywords.append(keyword)

    keyword_total = max(1, len(all_keywords))
    keyword_count = len(matched_keywords)
    score = 55 + round((keyword_count / keyword_total) * 30)
    if data['personal'].get('full_name') and data['personal'].get('email') and data['personal'].get('phone'):
        score += 8
    if data['job_requirements'].get('job_title') and data['job_requirements'].get('job_description'):
        score += 7
    if data.get('skill_categories'):
        score += 5
    if data.get('projects') or data.get('experience') or data.get('education'):
        score += 5
    score = min(score, 98)
    suggestions = [kw for kw in all_keywords if kw and not re.search(r'\b' + re.escape(kw) + r'\b', text)]
    return score, suggestions, keyword_count, keyword_total


def landing_page(request):
    content = _seed_default_content()
    context = {
        'content': content,
        'features': FeatureItem.objects.all(),
        'steps': ProcessStep.objects.all(),
        'benefits': BenefitItem.objects.all(),
        'testimonials': Testimonial.objects.all(),
        'faqs': FAQItem.objects.all(),
    }
    return render(request, 'cv_app/landing.html', context)


def builder(request):
    step = int(request.GET.get('step', 1))
    data = _get_builder_data(request)

    if request.method == 'POST':
        step = int(request.POST.get('step', step))

        if step == 1:
            data['personal'] = {
                'full_name': request.POST.get('full_name', '').strip(),
                'email': request.POST.get('email', '').strip(),
                'phone': request.POST.get('phone', '').strip(),
                'linkedin': request.POST.get('linkedin', '').strip(),
                'github': request.POST.get('github', '').strip(),
                'portfolio': request.POST.get('portfolio', '').strip(),
                'location': request.POST.get('location', '').strip(),
            }

        elif step == 2:
            data['job_requirements'] = {
                'job_title': request.POST.get('job_title', '').strip(),
                'role': request.POST.get('role', '').strip(),
                'job_description': request.POST.get('job_description', '').strip(),
                'keywords': request.POST.get('keywords', '').strip(),
                'required_skills': request.POST.get('required_skills', '').strip(),
            }
            data['skill_categories'] = _extract_skill_categories(
                data['job_requirements'].get('job_description', ''),
                data['job_requirements'].get('required_skills', ''),
                data['job_requirements'].get('required_skills', ''),
                data['job_requirements'].get('keywords', ''),
            )

        elif step == 3:
            skill_categories = []
            category_names = request.POST.getlist('skill_category[]')
            category_items = request.POST.getlist('skill_items[]')
            for cat, items in zip(category_names, category_items):
                if cat.strip() or items.strip():
                    skill_categories.append({'category': cat.strip(), 'items': items.strip()})
            data['skill_categories'] = skill_categories if skill_categories else data.get('skill_categories', [])

        elif step == 4:
            project_titles = request.POST.getlist('project_title[]')
            project_categories = request.POST.getlist('project_category[]')
            project_descriptions = request.POST.getlist('project_description[]')
            source_code_links = request.POST.getlist('source_code_link[]')
            project_links = request.POST.getlist('project_link[]')

            data['projects'] = [
                {
                    'title': title.strip(),
                    'category': category.strip(),
                    'description': description.strip(),
                    'source_code_link': source_code.strip(),
                    'project_link': link.strip(),
                }
                for title, category, description, source_code, link in zip(
                    project_titles, project_categories, project_descriptions, source_code_links, project_links
                )
                if title.strip() or category.strip() or description.strip() or source_code.strip() or link.strip()
            ]

        elif step == 5:
            company_names = request.POST.getlist('experience_company[]')
            experience_roles = request.POST.getlist('experience_role[]')
            experience_periods = request.POST.getlist('experience_period[]')
            experience_responsibilities = request.POST.getlist('experience_responsibilities[]')

            if company_names or experience_roles:
                data['experience'] = [
                    {
                        'company': company.strip(),
                        'role': role.strip(),
                        'period': period.strip(),
                        'responsibilities': responsibilities.strip(),
                    }
                    for company, role, period, responsibilities in zip(
                        company_names, experience_roles, experience_periods, experience_responsibilities
                    )
                    if company.strip() or role.strip() or period.strip() or responsibilities.strip()
                ]

        elif step == 6:
            degrees = request.POST.getlist('education_degree[]')
            institutions = request.POST.getlist('education_institution[]')
            periods = request.POST.getlist('education_period[]')
            gpa_cgpas = request.POST.getlist('education_gpa_cgpa[]')
            gpa_maxes = request.POST.getlist('education_gpa_max[]')
            education_details = request.POST.getlist('education_details[]')

            data['education'] = [
                {
                    'degree': degree.strip(),
                    'institution': institution.strip(),
                    'period': period.strip(),
                    'gpa_cgpa': gpa_cgpa.strip(),
                    'gpa_max': gpa_max.strip(),
                    'details': details.strip(),
                }
                for degree, institution, period, gpa_cgpa, gpa_max, details in zip(
                    degrees, institutions, periods, gpa_cgpas, gpa_maxes, education_details
                )
                if degree.strip() or institution.strip() or period.strip() or gpa_cgpa.strip() or gpa_max.strip() or details.strip()
            ]

        request.session['builder_data'] = data

        if step < 7:
            return redirect(f'/create/?step={step + 1}')
        return redirect('/preview/')

    if not data.get('skill_categories') and data.get('job_requirements', {}).get('job_description'):
        data['skill_categories'] = _extract_skill_categories(
            data['job_requirements'].get('job_description', ''),
            data['job_requirements'].get('required_skills', ''),
            data['job_requirements'].get('required_skills', ''),
            data['job_requirements'].get('keywords', ''),
        )
        request.session['builder_data'] = data

    context = {
        'step': step,
        'data': data,
        'skill_categories': data.get('skill_categories', []),
        'default_skill_categories': DEFAULT_SKILL_CATEGORIES,
    }
    return render(request, 'cv_app/builder.html', context)


def preview(request):
    data = _get_builder_data(request)
    score, suggestions, keyword_count, keyword_total = _calculate_score(data)
    context = {
        'data': data,
        'score': score,
        'keyword_count': keyword_count,
        'keyword_total': keyword_total,
        'keyword_suggestions': suggestions[:6],
        'generated_summary': _generate_summary(data),
    }
    return render(request, 'cv_app/preview.html', context)


def download_pdf(request):
    data = _get_builder_data(request)
    generated_summary = _generate_summary(data)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, leading=24, textColor=colors.HexColor('#1d4ed8'))
    heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, leading=15, textColor=colors.HexColor('#111827'))
    subheading_style = ParagraphStyle('SubHeading', parent=styles['Heading3'], fontName='Helvetica-Bold', fontSize=11, leading=13, textColor=colors.HexColor('#1d4ed8'))
    body_style = ParagraphStyle('Body', parent=styles['BodyText'], fontName='Helvetica', fontSize=10, leading=13, textColor=colors.HexColor('#111827'))
    small_style = ParagraphStyle('Small', parent=styles['BodyText'], fontName='Helvetica', fontSize=9, leading=11, textColor=colors.HexColor('#4b5563'))
    bullet_style = ParagraphStyle('Bullet', parent=body_style, leftIndent=20, bulletIndent=10)

    story = []

    name = data['personal'].get('full_name') or 'Your Name'
    story.append(Paragraph(name, title_style))

    contact_parts = []
    if data['personal'].get('location'):
        contact_parts.append(data['personal'].get('location'))
    if data['personal'].get('email'):
        contact_parts.append(data['personal'].get('email'))
    if data['personal'].get('phone'):
        contact_parts.append(data['personal'].get('phone'))
    story.append(Paragraph(' | '.join(contact_parts), small_style))

    links = []
    if data['personal'].get('linkedin'):
        links.append(f"LinkedIn: {data['personal'].get('linkedin')}")
    if data['personal'].get('github'):
        links.append(f"GitHub: {data['personal'].get('github')}")
    if data['personal'].get('portfolio'):
        links.append(f"Portfolio: {data['personal'].get('portfolio')}")
    if links:
        story.append(Paragraph(' • '.join(links), small_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('PROFESSIONAL SUMMARY', heading_style))
    story.append(Paragraph(generated_summary, body_style))
    story.append(Spacer(1, 8))

    story.append(Paragraph('PROFESSIONAL SKILLS', heading_style))
    for category in data.get('skill_categories', []):
        if category.get('items'):
            story.append(Paragraph(f"<b>{category.get('category')}</b>: {category.get('items')}", body_style))
    if not data.get('skill_categories'):
        story.append(Paragraph('<b>Communication</b>: Communication, Problem Solving', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('PROJECTS', heading_style))
    for project in data.get('projects', []):
        if not project.get('title') and not project.get('description'):
            continue

        title_parts = [f"<b>{project.get('title') or 'Project'}</b>"]
        if project.get('source_code_link'):
            title_parts.append(" <b>[Source Code]</b>")
        story.append(Paragraph(''.join(title_parts), body_style))

        if project.get('description'):
            bullets = []
            for line in project.get('description').splitlines():
                line = line.strip()
                if line:
                    bullets.append(line)
            if not bullets:
                bullets = [project.get('description').strip()]

            story.append(
                ListFlowable(
                    [ListItem(Paragraph(f"• {b}", bullet_style), bulletColor=colors.HexColor('#2563eb')) for b in bullets],
                    bulletType='bullet',
                )
            )

        links = []
        if project.get('source_code_link'):
            links.append(f"[Source Code] {project.get('source_code_link')}")
        if project.get('project_link'):
            links.append(f"Project Link: {project.get('project_link')}")
        if links:
            story.append(Paragraph(' • '.join(links), small_style))

        story.append(Spacer(1, 4))

    if not data.get('projects'):
        story.append(Paragraph('No project details entered', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('EDUCATION', heading_style))

    for education in data.get('education', []):
        if not education.get('institution') and not education.get('degree'):
            continue
        story.append(Paragraph(f"<b>{education.get('degree') or 'Degree'}</b> — {education.get('institution') or 'Institution'} | <b>Graduated</b>: {education.get('period')}", body_style))
        if education.get('gpa_cgpa') or education.get('gpa_max'):
            gpa_text = f"CGPA/GPA: {education.get('gpa_cgpa')}"
            if education.get('gpa_max'):
                gpa_text += f" / {education.get('gpa_max')}"
            story.append(Paragraph(gpa_text, small_style))
        if education.get('details'):
            story.append(Paragraph(education.get('details'), body_style))
        story.append(Spacer(1, 4))

    if not data.get('education'):
        story.append(Paragraph('No education details entered', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('EXPERIENCE', heading_style))
    for experience in data.get('experience', []):
        if not experience.get('company') and not experience.get('role'):
            continue
        story.append(Paragraph(f"<b>{experience.get('role') or 'Role'}</b> @ {experience.get('company') or 'Company'}", body_style))
        if experience.get('period'):
            story.append(Paragraph(experience.get('period'), small_style))
        responsibilities = [item.strip() for item in experience.get('responsibilities', '').split('\n') if item.strip()]
        if responsibilities:
            story.append(
                ListFlowable(
                    [ListItem(Paragraph(f"• {item}", bullet_style), bulletColor=colors.HexColor('#2563eb')) for item in responsibilities],
                    bulletType='bullet',
                )
            )
        story.append(Spacer(1, 4))

    if not data.get('experience'):
        story.append(Paragraph('No experience details entered', body_style))

    doc.build(story)
    pdf_data = buffer.getvalue()
    response = HttpResponse(pdf_data, content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="cv_guide_cv.pdf"'
    return response


def download_docx(request):
    data = _get_builder_data(request)
    document = Document()

    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x14, 0x21, 0x3D)

    # Name
    heading = document.add_heading(data['personal'].get('full_name') or 'Your Name', level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(24)

    # Contact info
    contact_parts = []
    if data['personal'].get('location'):
        contact_parts.append(data['personal'].get('location'))
    if data['personal'].get('email'):
        contact_parts.append(data['personal'].get('email'))
    if data['personal'].get('phone'):
        contact_parts.append(data['personal'].get('phone'))
    p = document.add_paragraph(' | '.join(contact_parts))
    p.style.font.size = Pt(10)
    p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    links = []
    if data['personal'].get('linkedin'):
        links.append(f"LinkedIn: {data['personal'].get('linkedin')}")
    if data['personal'].get('github'):
        links.append(f"GitHub: {data['personal'].get('github')}")
    if data['personal'].get('portfolio'):
        links.append(f"Portfolio: {data['personal'].get('portfolio')}")
    if links:
        p = document.add_paragraph(' • '.join(links))
        p.style.font.size = Pt(10)
        p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    document.add_paragraph()  # spacer

    # Professional Summary
    heading = document.add_heading('PROFESSIONAL SUMMARY', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    document.add_paragraph(_generate_summary(data))

    # Professional Skills
    heading = document.add_heading('PROFESSIONAL SKILLS', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for category in data.get('skill_categories', []):
        if category.get('items'):
            p = document.add_paragraph()
            run = p.add_run(f"- {category.get('category')}: ")
            run.bold = True
            p.add_run(category.get('items'))

    # Projects
    heading = document.add_heading('PROJECTS', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for project in data.get('projects', []):
        if project.get('title') or project.get('description'):
            p = document.add_paragraph()
            run = p.add_run(project.get('title') or 'Project')
            run.bold = True
            if project.get('source_code_link'):
                run2 = p.add_run(' [Source Code]')
                run2.bold = True

            if project.get('description'):
                lines = [ln.strip() for ln in project.get('description').splitlines() if ln.strip()]
                if not lines:
                    lines = [project.get('description').strip()]
                for b in lines:
                    document.add_paragraph(b, style='List Bullet')

            link_parts = []
            if project.get('source_code_link'):
                link_parts.append(f"[Source Code] {project.get('source_code_link')}")
            if project.get('project_link'):
                link_parts.append(f"Project Link: {project.get('project_link')}")
            if link_parts:
                p = document.add_paragraph(' • '.join(link_parts))
                p.style.font.size = Pt(9)
                p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Education
    heading = document.add_heading('EDUCATION', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for education in data.get('education', []):
        if education.get('institution') or education.get('degree'):
            degree = education.get('degree') or 'Degree'
            inst = education.get('institution') or 'Institution'
            year = education.get('period') or ''
            p = document.add_paragraph()
            run = p.add_run(f"{degree} | {inst} | Graduated: {year}")
            run.bold = True

            gpa = education.get('gpa_cgpa')
            gmax = education.get('gpa_max')
            if gpa or gmax:
                if gmax:
                    document.add_paragraph(f"CGPA/GPA: {gpa} / {gmax}")
                else:
                    document.add_paragraph(f"CGPA/GPA: {gpa}")

            if education.get('details'):
                document.add_paragraph(education.get('details'))

    # Experience
    heading = document.add_heading('EXPERIENCE', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for experience in data.get('experience', []):
        if experience.get('company') or experience.get('role'):
            p = document.add_paragraph()
            run = p.add_run(f"{experience.get('role') or 'Role'} @ {experience.get('company') or 'Company'}")
            run.bold = True
            if experience.get('period'):
                document.add_paragraph(experience.get('period'))
            for responsibility in [item.strip() for item in experience.get('responsibilities', '').split('\n') if item.strip()]:
                document.add_paragraph(responsibility, style='List Bullet')

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="cv_guide_cv.docx"'
    document.save(response)
    return response


@login_required(login_url='/cv-admin/login/')
def cv_admin_dashboard(request):
    content = LandingPageContent.objects.first() or LandingPageContent.objects.create()
    FeatureFormSet = modelformset_factory(FeatureItem, form=FeatureItemForm, extra=1)
    ProcessFormSet = modelformset_factory(ProcessStep, form=ProcessStepForm, extra=1)
    BenefitFormSet = modelformset_factory(BenefitItem, form=BenefitItemForm, extra=1)
    TestimonialFormSet = modelformset_factory(Testimonial, form=TestimonialForm, extra=1)
    FAQFormSet = modelformset_factory(FAQItem, form=FAQItemForm, extra=1)

    if request.method == 'POST':
        main_form = LandingPageContentForm(request.POST, instance=content)
        features_formset = FeatureFormSet(request.POST, queryset=FeatureItem.objects.all())
        steps_formset = ProcessFormSet(request.POST, queryset=ProcessStep.objects.all())
        benefits_formset = BenefitFormSet(request.POST, queryset=BenefitItem.objects.all())
        testimonials_formset = TestimonialFormSet(request.POST, queryset=Testimonial.objects.all())
        faqs_formset = FAQFormSet(request.POST, queryset=FAQItem.objects.all())
        if all([
            main_form.is_valid(),
            features_formset.is_valid(),
            steps_formset.is_valid(),
            benefits_formset.is_valid(),
            testimonials_formset.is_valid(),
            faqs_formset.is_valid(),
        ]):
            main_form.save()
            features_formset.save()
            steps_formset.save()
            benefits_formset.save()
            testimonials_formset.save()
            faqs_formset.save()
            messages.success(request, 'Landing page content updated.')
            return redirect('cv_admin_dashboard')
    else:
        main_form = LandingPageContentForm(instance=content)
        features_formset = FeatureFormSet(queryset=FeatureItem.objects.all())
        steps_formset = ProcessFormSet(queryset=ProcessStep.objects.all())
        benefits_formset = BenefitFormSet(queryset=BenefitItem.objects.all())
        testimonials_formset = TestimonialFormSet(queryset=Testimonial.objects.all())
        faqs_formset = FAQFormSet(queryset=FAQItem.objects.all())

    context = {
        'main_form': main_form,
        'features_formset': features_formset,
        'steps_formset': steps_formset,
        'benefits_formset': benefits_formset,
        'testimonials_formset': testimonials_formset,
        'faqs_formset': faqs_formset,
    }
    return render(request, 'cv_app/cv_admin.html', context)


def cv_admin_login(request):
    if request.user.is_authenticated:
        return redirect('cv_admin_dashboard')
    error = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '').strip()
        user = authenticate(request, username=username, password=password)
        if user is not None and user.is_staff:
            login(request, user)
            return redirect('cv_admin_dashboard')
        error = 'Please use a valid staff account.'
    return render(request, 'cv_app/admin_login.html', {'error': error})


def cv_admin_logout(request):
    logout(request)
    return redirect('cv_admin_login')