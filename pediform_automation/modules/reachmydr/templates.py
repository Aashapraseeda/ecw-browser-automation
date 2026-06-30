"""
modules/reachmydr/templates.py
-------------------------------
Message templates for ReachMyDr reminders.
{form_link} is replaced at send time with the actual URL from settings.
"""

REMINDER_MESSAGE = (
    "Hello! Your child has an upcoming well-child visit. "
    "Please complete the required health form before your appointment: {form_link}"
)

MORNING_REMINDER_MESSAGE = (
    "Good morning! Your child's appointment is today. "
    "Please fill out the health form now: {form_link}"
)

EVENING_REMINDER_MESSAGE = (
    "Reminder: Please complete your child's health form before tomorrow's appointment: {form_link}"
)
