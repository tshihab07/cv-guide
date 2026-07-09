from django.contrib import admin

from .models import (
    BenefitItem,
    FAQItem,
    FeatureItem,
    LandingPageContent,
    ProcessStep,
    Testimonial,
)

admin.site.register(LandingPageContent)
admin.site.register(FeatureItem)
admin.site.register(ProcessStep)
admin.site.register(BenefitItem)
admin.site.register(Testimonial)
admin.site.register(FAQItem)
