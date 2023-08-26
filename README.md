gcloud builds submit --tag gcr.io/aman-fs/<AppName>  --project=<ProjectName>

gcloud run deploy --image gcr.io/aman-fs/<AppName> --platform managed  --project=<ProjectName> --allow-unauthenticated