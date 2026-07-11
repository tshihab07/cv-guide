from django import forms

from .models import (
    BenefitItem,
    FAQItem,
    FeatureItem,
    LandingPageContent,
    ProcessStep,
    Testimonial,
)


class CertificationForm(forms.Form):
    name = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'AWS Certified Solutions Architect'}))
    issuer = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Amazon Web Services'}))
    date_obtained = forms.CharField(max_length=100, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Jan 2024'}))
    expiration_date = forms.CharField(max_length=100, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Jan 2027 (optional)'}))
    credential_id = forms.CharField(max_length=100, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'ABC123XYZ (optional)'}))
    credential_url = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://www.credly.com/badges/... (optional)'}))


class HonorAwardForm(forms.Form):
    title = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': "Dean's List / Best Paper Award"}))
    issuer = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'University Name / Conference Name'}))
    date_received = forms.CharField(max_length=100, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'May 2023'}))
    description = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 2, 'placeholder': 'Brief description of the honor/award (optional)'}))


class LandingPageContentForm(forms.ModelForm):
    class Meta:
        model = LandingPageContent
        fields = [
            'hero_title',
            'hero_subtitle',
            'hero_cta_text',
            'safety_notice',
            'features_title',
            'features_subtitle',
            'process_title',
            'process_subtitle',
            'benefits_title',
            'benefits_subtitle',
            'testimonials_title',
            'faq_title',
            'footer_text',
        ]
        widgets = {
            'hero_title': forms.TextInput(attrs={'class': 'form-control'}),
            'hero_subtitle': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'hero_cta_text': forms.TextInput(attrs={'class': 'form-control'}),
            'safety_notice': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'features_title': forms.TextInput(attrs={'class': 'form-control'}),
            'features_subtitle': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'process_title': forms.TextInput(attrs={'class': 'form-control'}),
            'process_subtitle': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'benefits_title': forms.TextInput(attrs={'class': 'form-control'}),
            'benefits_subtitle': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'testimonials_title': forms.TextInput(attrs={'class': 'form-control'}),
            'faq_title': forms.TextInput(attrs={'class': 'form-control'}),
            'footer_text': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }


class FeatureItemForm(forms.ModelForm):
    class Meta:
        model = FeatureItem
        fields = ['title', 'description', 'icon', 'order']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'icon': forms.TextInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class ProcessStepForm(forms.ModelForm):
    class Meta:
        model = ProcessStep
        fields = ['title', 'description', 'order']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class BenefitItemForm(forms.ModelForm):
    class Meta:
        model = BenefitItem
        fields = ['title', 'description', 'icon', 'order']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'icon': forms.TextInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class TestimonialForm(forms.ModelForm):
    class Meta:
        model = Testimonial
        fields = ['quote', 'author', 'role', 'order']
        widgets = {
            'quote': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'author': forms.TextInput(attrs={'class': 'form-control'}),
            'role': forms.TextInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class FAQItemForm(forms.ModelForm):
    class Meta:
        model = FAQItem
        fields = ['question', 'answer', 'order']
        widgets = {
            'question': forms.TextInput(attrs={'class': 'form-control'}),
            'answer': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class PersonalInfoForm(forms.Form):
    full_name = forms.CharField(max_length=200, required=True, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Md Tushar Shihab'}))
    email = forms.EmailField(required=True, widget=forms.EmailInput(attrs={'class': 'form-control', 'placeholder': 'tushar@example.com'}))
    phone = forms.CharField(max_length=50, required=True, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': '+880 1XXXXXXXXX'}))
    location = forms.CharField(max_length=200, required=True, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Dhaka, Bangladesh'}))
    linkedin = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://linkedin.com/in/username'}))
    github = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://github.com/username'}))
    portfolio = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://yourportfolio.com'}))


class JobRequirementsForm(forms.Form):
    job_title = forms.CharField(max_length=200, required=True, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Machine Learning Engineer'}))
    role = forms.CharField(max_length=200, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'ML Engineer / Data Scientist'}))
    job_description = forms.CharField(required=True, widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 6, 'placeholder': 'Paste the full job description here...'}))
    keywords = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Python, TensorFlow, AWS, Docker, MLOps'}))
    required_skills = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Python, PyTorch, Kubernetes, CI/CD'}))


class SkillCategoryForm(forms.Form):
    category = forms.ChoiceField(choices=[
        ('Programming Languages', 'Programming Languages'),
        ('Machine Learning & AI', 'Machine Learning & AI'),
        ('Data Analysis & Visualization', 'Data Analysis & Visualization'),
        ('Databases', 'Databases'),
        ('Backend Development', 'Backend Development'),
        ('Frontend Development', 'Frontend Development'),
        ('Cloud & DevOps', 'Cloud & DevOps'),
        ('Web Scraping & Automation', 'Web Scraping & Automation'),
        ('Testing & QA', 'Testing & QA'),
        ('Project Management & Soft Skills', 'Project Management & Soft Skills'),
    ], widget=forms.Select(attrs={'class': 'form-control'}))
    items = forms.CharField(widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 3, 'placeholder': 'Python, C++, JavaScript, SQL'}))


class ProjectForm(forms.Form):
    title = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Intelligent Loan Approval Prediction System'}))
    category = forms.ChoiceField(choices=[
        ('Machine Learning', 'Machine Learning'),
        ('Python', 'Python'),
        ('Django', 'Django'),
        ('Flask', 'Flask'),
        ('FastAPI', 'FastAPI'),
        ('ReactJS', 'ReactJS'),
        ('AngularJS', 'AngularJS'),
        ('VueJS', 'VueJS'),
        ('Event Management', 'Event Management'),
        ('Data Analysis', 'Data Analysis'),
        ('Web Scraping', 'Web Scraping'),
        ('Automation', 'Automation'),
        ('Backend Development', 'Backend Development'),
        ('Frontend Development', 'Frontend Development'),
        ('Full Stack', 'Full Stack'),
        ('Mobile Development', 'Mobile Development'),
        ('DevOps', 'DevOps'),
        ('Cloud Computing', 'Cloud Computing'),
        ('Cybersecurity', 'Cybersecurity'),
        ('Blockchain', 'Blockchain'),
        ('IoT', 'IoT'),
        ('Game Development', 'Game Development'),
        ('Desktop Application', 'Desktop Application'),
        ('Research', 'Research'),
        ('Academic Project', 'Academic Project'),
        ('Other', 'Other'),
    ], widget=forms.Select(attrs={'class': 'form-control'}))
    description = forms.CharField(widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 4, 'placeholder': 'Performed EDA and feature engineering on financial datasets and by using advanced machine learning models(Decision Tree, GBM, etc), the system captures complex patterns\nReduced predicted default risk by 86% using XGBoost with optuna optimization\nEnsure fair and consistent evaluation\nCut manual screening time by 67% through automated ML pipeline'}))
    source_code_link = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://github.com/...'}))
    project_link = forms.URLField(required=False, widget=forms.URLInput(attrs={'class': 'form-control', 'placeholder': 'https://your-project-url.com'}))


class EducationForm(forms.Form):
    degree = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'B.Sc in Computer Science and Engineering'}))
    institution = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'National University, Bangladesh'}))
    period = forms.CharField(max_length=100, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': '2020 - 2024'}))
    gpa_cgpa = forms.CharField(max_length=20, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': '3.75'}))
    gpa_max = forms.CharField(max_length=20, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': '4.00'}))
    details = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 3, 'placeholder': 'Relevant coursework: Data Structures, Algorithms, Machine Learning, Database Systems'}))


class ExperienceForm(forms.Form):
    company = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'ABC Technologies Ltd.'}))
    role = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Software Engineer'}))
    period = forms.CharField(max_length=100, required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Jan 2023 - Present'}))
    responsibilities = forms.CharField(widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 4, 'placeholder': 'Developed and maintained REST APIs using Django REST Framework\nImplemented CI/CD pipelines using GitHub Actions\nCollaborated with cross-functional teams to deliver features on time'}))