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
