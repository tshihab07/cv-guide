from django.db import models


class LandingPageContent(models.Model):
    hero_title = models.CharField(max_length=200, default='Build ATS-Friendly CVs in Minutes')
    hero_subtitle = models.TextField(default='Tailored to the job requirements you care about.')
    hero_cta_text = models.CharField(max_length=80, default='Create Your CV')
    safety_notice = models.TextField(default='Your information stays in the current session and is not stored long term.')
    features_title = models.CharField(max_length=200, default='Features')
    features_subtitle = models.TextField(default='Everything you need to build a polished CV quickly.')
    process_title = models.CharField(max_length=200, default='How it works')
    process_subtitle = models.TextField(default='A simple five-step flow from input to download.')
    benefits_title = models.CharField(max_length=200, default='What you get')
    benefits_subtitle = models.TextField(default='A recruiter-friendly CV with relevant keywords and a clear layout.')
    testimonials_title = models.CharField(max_length=200, default='What users say')
    faq_title = models.CharField(max_length=200, default='Frequently asked questions')
    footer_text = models.TextField(default='© 2026 CVGuide. Built for ATS-friendly CV creation.')

    class Meta:
        verbose_name_plural = 'Landing Page Content'

    def __str__(self):
        return 'Landing Page Content'


class FeatureItem(models.Model):
    title = models.CharField(max_length=120)
    description = models.TextField()
    icon = models.CharField(max_length=20, default='⚡')
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.title


class ProcessStep(models.Model):
    title = models.CharField(max_length=120)
    description = models.TextField()
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.title


class BenefitItem(models.Model):
    title = models.CharField(max_length=120)
    description = models.TextField()
    icon = models.CharField(max_length=20, default='🎯')
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.title


class Testimonial(models.Model):
    quote = models.TextField()
    author = models.CharField(max_length=120)
    role = models.CharField(max_length=120)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.author


class FAQItem(models.Model):
    question = models.CharField(max_length=200)
    answer = models.TextField()
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.question


class SkillCategory(models.Model):
    CATEGORY_CHOICES = [
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
    ]
    category = models.CharField(max_length=100, choices=CATEGORY_CHOICES)
    items = models.TextField(help_text="Comma-separated skills")
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']
        verbose_name_plural = 'Skill Categories'

    def __str__(self):
        return self.category


class Project(models.Model):
    CATEGORY_CHOICES = [
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
    ]
    title = models.CharField(max_length=200)
    category = models.CharField(max_length=50, choices=CATEGORY_CHOICES)
    description = models.TextField(help_text="Use bullet points (one per line)")
    source_code_link = models.URLField(blank=True, null=True)
    project_link = models.URLField(blank=True, null=True)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.title


class Education(models.Model):
    degree = models.CharField(max_length=200)
    institution = models.CharField(max_length=200)
    period = models.CharField(max_length=100, help_text="e.g., 2020 - 2024")
    gpa_cgpa = models.CharField(max_length=20, blank=True, null=True)
    gpa_max = models.CharField(max_length=20, blank=True, null=True)
    details = models.TextField(blank=True, null=True)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']
        verbose_name_plural = 'Education'

    def __str__(self):
        return f"{self.degree} - {self.institution}"


class Experience(models.Model):
    company = models.CharField(max_length=200)
    role = models.CharField(max_length=200)
    period = models.CharField(max_length=100, blank=True, null=True)
    responsibilities = models.TextField(help_text="One responsibility per line")
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return f"{self.role} @ {self.company}"


class PersonalInfo(models.Model):
    full_name = models.CharField(max_length=200)
    email = models.EmailField()
    phone = models.CharField(max_length=50)
    location = models.CharField(max_length=200)
    linkedin = models.URLField(blank=True, null=True)
    github = models.URLField(blank=True, null=True)
    portfolio = models.URLField(blank=True, null=True)

    def __str__(self):
        return self.full_name


class JobRequirements(models.Model):
    job_title = models.CharField(max_length=200)
    role = models.CharField(max_length=200, blank=True, null=True)
    job_description = models.TextField()
    keywords = models.TextField(blank=True, null=True, help_text="Comma-separated keywords")
    required_skills = models.TextField(blank=True, null=True, help_text="Comma-separated required skills")

    def __str__(self):
        return self.job_title


class Certification(models.Model):
    name = models.CharField(max_length=200)
    issuing_organization = models.CharField(max_length=200)
    issue_date = models.CharField(max_length=100, help_text="e.g., Jan 2024")
    expiration_date = models.CharField(max_length=100, blank=True, null=True, help_text="e.g., Jan 2027 or 'Does not expire'")
    credential_id = models.CharField(max_length=200, blank=True, null=True)
    credential_url = models.URLField(blank=True, null=True)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']
        verbose_name_plural = 'Certifications'

    def __str__(self):
        return f"{self.name} - {self.issuing_organization}"


class HonorAward(models.Model):
    title = models.CharField(max_length=200)
    issuer = models.CharField(max_length=200)
    date_received = models.CharField(max_length=100, help_text="e.g., June 2023")
    description = models.TextField(blank=True, null=True, help_text="Brief description of the honor/award")
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']
        verbose_name = 'Honor & Award'
        verbose_name_plural = 'Honors & Awards'

    def __str__(self):
        return self.title