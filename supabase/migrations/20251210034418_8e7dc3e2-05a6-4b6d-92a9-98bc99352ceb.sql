
-- Update intent registrations status to approved for Sepik River Transport Ltd and Island Fisheries PNG Ltd
UPDATE intent_registrations 
SET status = 'approved', updated_at = now()
WHERE id IN (
  'ea93f164-0e39-42fe-be43-f2f7b1d4d98a',  -- Island Fisheries PNG Ltd - Alotau Commercial Fishing Marina Development
  '39a23fd2-10e0-4d98-ba70-d0ccf42b5225'   -- Sepik River Transport Ltd - Angoram River Transport Hub Development
);
