# Quiz Processing System

Quiz Processing System extracts topic scores from quiz PDFs and exports CSV or Excel results. It supports optional local gradebooks and an optional Google Sheets gradebook created by the application.

Google Sheets developer setup and the complete test checklist are documented in [`developer_implementation.md`](developer_implementation.md). Do not use or distribute a service-account key; the application uses per-teacher Desktop OAuth with the `drive.file` scope.
