# Onboarding Flow

This simple LAMP stack onboarding demo uses PHP and jQuery UI to capture
seller information before allowing access to a dashboard. Each step updates
a progress bar and persists data in MySQL.

## Setup

1. Create a database named `onboarding_db` and run the SQL in
   `onboarding/onboarding.sql` to create the `seller_onboarding` table.
2. Configure your Apache/PHP environment so the `onboarding` directory is
   accessible via the web server.
3. Adjust database credentials in `onboarding/db.php` if needed.

## Usage

Navigate to `onboarding/step1.php` to start. Progress is saved through
sessions and written to the `seller_onboarding` table.
