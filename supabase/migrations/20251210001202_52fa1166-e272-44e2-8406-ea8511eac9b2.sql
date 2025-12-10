-- Add project_title column to intent_registrations table
ALTER TABLE public.intent_registrations 
ADD COLUMN IF NOT EXISTS project_title text;

-- Add project_title column to intent_registration_drafts table
ALTER TABLE public.intent_registration_drafts 
ADD COLUMN IF NOT EXISTS project_title text;

-- Add comment for documentation
COMMENT ON COLUMN public.intent_registrations.project_title IS 'Title of the project for the intent registration';
COMMENT ON COLUMN public.intent_registration_drafts.project_title IS 'Title of the project for the intent registration draft';