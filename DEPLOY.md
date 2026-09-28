# Enhanced deployment – Uganda Report Card System

## Render environment variables

Keep the existing `DATABASE_URL` and `SECRET_KEY`.

Add these variables in Render > Environment:

- `SCHOOL_NAME` = your school's name
- `SCHOOL_MOTTO` = your school motto
- `SCHOOL_ADDRESS` = school address
- `SCHOOL_PHONE` = school phone
- `SCHOOL_EMAIL` = school email
- `SCHOOL_DISTRICT` = district
- `ADMIN_NAME` = name of school administrator
- `ADMIN_EMAIL` = administrator email
- `ADMIN_PASSWORD` = a strong administrator password

Do not put real passwords into GitHub.

## What was added

1. PostgreSQL-backed users and school records.
2. Administrator and teacher roles.
3. Teacher mark-entry screen.
4. Central `mark_entries` table.
5. Administrator dashboard showing recent mark updates.
6. Audit log for login, mark entry, teacher creation and subscription activation.
7. UGX 150,000 term subscription records.
8. Automatic expiry after 90 days.
9. Existing Excel/PDF report-card workflow retained.

## Subscription payment

This first production-ready stage records a verified payment/mobile-money reference and lets the administrator activate the 90-day subscription. It does not pretend to process MTN/Airtel money automatically.

A payment gateway can be connected in the next stage once the preferred payment provider and merchant account are selected.
