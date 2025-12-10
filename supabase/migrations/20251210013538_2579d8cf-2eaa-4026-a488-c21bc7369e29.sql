-- Update intent registration drafts with missing project titles based on activity descriptions

UPDATE intent_registration_drafts 
SET project_title = 'Markham Valley Oil & Gas Exploration Project', updated_at = now()
WHERE id = '3c85574f-7b84-4c61-8b68-bc3c9db75751';

UPDATE intent_registration_drafts 
SET project_title = 'Port Moresby 150MW Gas Power Station', updated_at = now()
WHERE id = '3dc03a5c-e1c6-4d82-bc44-64dd35c47032';