from django import template

register = template.Library()


@register.filter
def split_items(value):
    if not value:
        return []
    if isinstance(value, str):
        return [item.strip() for item in value.split(',') if item.strip()]
    return value


@register.filter
def trim(value):
    if value is None:
        return ''
    return str(value).strip()
