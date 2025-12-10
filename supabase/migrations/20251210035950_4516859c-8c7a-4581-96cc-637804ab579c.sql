-- Add timestamp columns for legal declarations
ALTER TABLE public.permit_applications 
ADD COLUMN IF NOT EXISTS legal_declaration_accepted_at TIMESTAMP WITH TIME ZONE,
ADD COLUMN IF NOT EXISTS compliance_commitment_accepted_at TIMESTAMP WITH TIME ZONE;

-- Add comment for documentation
COMMENT ON COLUMN public.permit_applications.legal_declaration_accepted_at IS 'Timestamp when applicant accepted the legal declaration';
COMMENT ON COLUMN public.permit_applications.compliance_commitment_accepted_at IS 'Timestamp when applicant accepted the compliance commitment';