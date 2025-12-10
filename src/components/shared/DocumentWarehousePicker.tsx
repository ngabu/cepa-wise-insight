import React, { useState, useMemo, useRef } from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogFooter } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Badge } from '@/components/ui/badge';
import { Card, CardContent } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { ScrollArea } from '@/components/ui/scroll-area';
import { FileText, Search, FolderOpen, Upload, Loader2, Check, Library } from 'lucide-react';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import { supabase } from '@/integrations/supabase/client';
import { useAuth } from '@/contexts/AuthContext';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { useDocuments } from '@/hooks/useDocuments';

interface DocumentFile {
  id: string;
  filename: string;
  file_path: string;
  file_size: number;
  mime_type: string;
  uploaded_at: string;
  document_type?: string;
}

interface DocumentWarehousePickerProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSelect: (documents: DocumentFile[]) => void;
  multiSelect?: boolean;
  title?: string;
  description?: string;
  acceptedTypes?: string[];
}

const categories = [
  { id: 'all', name: 'All Documents' },
  { id: 'project_details', name: 'Project Details' },
  { id: 'environmental_assessment', name: 'Environmental Assessment' },
  { id: 'monitoring', name: 'Monitoring' },
  { id: 'compliance_reports', name: 'Compliance Reports' },
  { id: 'inspection_reports', name: 'Inspection Reports' },
  { id: 'notices', name: 'Notices' },
  { id: 'safety', name: 'Safety' },
];

const formatFileSize = (bytes: number) => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

const getCategoryName = (categoryId: string) => {
  const category = categories.find(c => c.id === categoryId);
  return category ? category.name : 'General Document';
};

export function DocumentWarehousePicker({
  open,
  onOpenChange,
  onSelect,
  multiSelect = true,
  title = "Select Documents from Warehouse",
  description = "Choose documents from your Document Management to attach. These documents can be reused across applications.",
  acceptedTypes = [],
}: DocumentWarehousePickerProps) {
  const { user } = useAuth();
  const queryClient = useQueryClient();
  const { uploadDocumentToWarehouse } = useDocuments();
  const [selectedCategory, setSelectedCategory] = useState('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDocuments, setSelectedDocuments] = useState<DocumentFile[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper function to upload document with category
  const uploadDocument = async (file: File, category?: string) => {
    return uploadDocumentToWarehouse(file, category);
  };

  const { data: documents = [], isLoading } = useQuery({
    queryKey: ['warehouse-documents', user?.id],
    queryFn: async () => {
      if (!user) return [];
      const { data, error } = await supabase
        .from('documents')
        .select('*')
        .eq('user_id', user.id)
        .order('uploaded_at', { ascending: false });

      if (error) throw error;
      return data as DocumentFile[];
    },
    enabled: !!user && open,
  });

  const filteredDocuments = useMemo(() => {
    return documents.filter(doc => {
      const matchesCategory = selectedCategory === 'all' || doc.document_type === selectedCategory;
      const matchesSearch = doc.filename.toLowerCase().includes(searchQuery.toLowerCase());
      const matchesType = acceptedTypes.length === 0 || 
        acceptedTypes.some(type => doc.mime_type?.includes(type) || doc.filename.endsWith(type.replace('*', '')));
      return matchesCategory && matchesSearch && matchesType;
    });
  }, [documents, selectedCategory, searchQuery, acceptedTypes]);

  const categoryCounts = useMemo(() => {
    const counts: Record<string, number> = { all: documents.length };
    categories.forEach(cat => {
      if (cat.id !== 'all') {
        counts[cat.id] = documents.filter(f => f.document_type === cat.id).length;
      }
    });
    return counts;
  }, [documents]);

  const handleDocumentToggle = (doc: DocumentFile) => {
    if (multiSelect) {
      setSelectedDocuments(prev => {
        const isSelected = prev.some(d => d.id === doc.id);
        if (isSelected) {
          return prev.filter(d => d.id !== doc.id);
        }
        return [...prev, doc];
      });
    } else {
      setSelectedDocuments([doc]);
    }
  };

  const handleConfirm = () => {
    onSelect(selectedDocuments);
    setSelectedDocuments([]);
    onOpenChange(false);
  };

  const handleClose = () => {
    setSelectedDocuments([]);
    onOpenChange(false);
  };

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    setIsUploading(true);
    try {
      const uploadedDocs: DocumentFile[] = [];
      
      // Determine the category to use - use selected category if not 'all', otherwise undefined
      const categoryToUse = selectedCategory !== 'all' ? selectedCategory : undefined;
      
      for (const file of Array.from(files)) {
        const result = await uploadDocument(file, categoryToUse);
        if (result) {
          uploadedDocs.push({
            id: result.id,
            filename: result.filename,
            file_path: result.file_path,
            file_size: result.file_size || file.size,
            mime_type: result.mime_type || file.type,
            uploaded_at: result.uploaded_at,
            document_type: result.document_type,
          });
        }
      }

      // Refresh the documents list
      await queryClient.invalidateQueries({ queryKey: ['warehouse-documents', user?.id] });
      
      // Auto-select the newly uploaded documents
      if (uploadedDocs.length > 0) {
        if (multiSelect) {
          setSelectedDocuments(prev => [...prev, ...uploadedDocs]);
        } else {
          setSelectedDocuments([uploadedDocs[0]]);
        }
        const categoryName = categoryToUse ? getCategoryName(categoryToUse) : 'warehouse';
        toast.success(`${uploadedDocs.length} document${uploadedDocs.length > 1 ? 's' : ''} uploaded to ${categoryName}`);
      }
    } catch (error: any) {
      console.error('Failed to upload document:', error);
      toast.error('Failed to upload document: ' + error.message);
    } finally {
      setIsUploading(false);
      // Reset file input
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  const isSelected = (docId: string) => selectedDocuments.some(d => d.id === docId);

  // Get accepted file types for input
  const acceptString = acceptedTypes.length > 0 
    ? acceptedTypes.map(t => t.startsWith('.') ? t : `.${t}`).join(',')
    : '.pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg';

  return (
    <Dialog open={open} onOpenChange={handleClose}>
      <DialogContent className="max-w-4xl max-h-[85vh] flex flex-col">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Library className="w-5 h-5 text-primary" />
            {title}
          </DialogTitle>
          <DialogDescription>{description}</DialogDescription>
        </DialogHeader>

        <div className="flex-1 flex gap-4 min-h-0 overflow-hidden">
          {/* Categories Sidebar */}
          <div className="w-56 flex-shrink-0 border-r pr-4">
            <h4 className="font-semibold text-sm text-foreground mb-3">Categories</h4>
            <ScrollArea className="h-[400px]">
              <div className="space-y-1">
                {categories.map((category) => (
                  <button
                    key={category.id}
                    onClick={() => setSelectedCategory(category.id)}
                    className={cn(
                      "w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm transition-colors",
                      selectedCategory === category.id
                        ? "bg-primary/10 text-primary font-medium"
                        : "text-muted-foreground hover:bg-muted hover:text-foreground"
                    )}
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <FolderOpen className="w-4 h-4 flex-shrink-0" />
                      <span className="truncate text-xs">{category.name}</span>
                    </div>
                    <Badge variant="secondary" className="text-xs px-1.5 py-0 flex-shrink-0 ml-1">
                      {categoryCounts[category.id] || 0}
                    </Badge>
                  </button>
                ))}
              </div>
            </ScrollArea>
          </div>

          {/* Documents List */}
          <div className="flex-1 flex flex-col min-h-0">
            {/* Search and Upload Row */}
            <div className="flex gap-2 mb-4">
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
                <Input
                  placeholder="Search documents..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-10"
                />
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept={acceptString}
                onChange={handleFileUpload}
                className="hidden"
                multiple={multiSelect}
              />
              <Button
                variant="outline"
                onClick={handleUploadClick}
                disabled={isUploading}
                className="flex-shrink-0"
              >
                {isUploading ? (
                  <>
                    <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                    Uploading...
                  </>
                ) : (
                  <>
                    <Upload className="w-4 h-4 mr-2" />
                    Upload New
                  </>
                )}
              </Button>
            </div>

            {/* Document List */}
            <ScrollArea className="flex-1">
              {isLoading ? (
                <div className="flex items-center justify-center py-12">
                  <Loader2 className="w-6 h-6 animate-spin text-primary" />
                  <span className="ml-2 text-muted-foreground">Loading documents...</span>
                </div>
              ) : filteredDocuments.length === 0 ? (
                <div className="text-center py-12">
                  <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mx-auto mb-4">
                    <FileText className="w-8 h-8 text-muted-foreground" />
                  </div>
                  <h3 className="text-lg font-medium text-foreground mb-2">No documents found</h3>
                  <p className="text-sm text-muted-foreground mb-4">
                    {documents.length === 0 
                      ? "Your document warehouse is empty. Upload a new document to get started." 
                      : "Try adjusting your search or category filter"}
                  </p>
                  <Button onClick={handleUploadClick} disabled={isUploading}>
                    {isUploading ? (
                      <>
                        <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                        Uploading...
                      </>
                    ) : (
                      <>
                        <Upload className="w-4 h-4 mr-2" />
                        Upload New Document
                      </>
                    )}
                  </Button>
                </div>
              ) : (
                <div className="space-y-2 pr-2">
                  {filteredDocuments.map((doc) => (
                    <Card 
                      key={doc.id} 
                      className={cn(
                        "cursor-pointer transition-all border",
                        isSelected(doc.id) 
                          ? "border-primary bg-primary/5 ring-1 ring-primary" 
                          : "hover:border-primary/30"
                      )}
                      onClick={() => handleDocumentToggle(doc)}
                    >
                      <CardContent className="p-3">
                        <div className="flex items-center gap-3">
                          <div className={cn(
                            "w-8 h-8 rounded-lg flex items-center justify-center flex-shrink-0",
                            isSelected(doc.id) ? "bg-primary" : "bg-blue-50 dark:bg-blue-950"
                          )}>
                            {isSelected(doc.id) ? (
                              <Check className="w-4 h-4 text-primary-foreground" />
                            ) : (
                              <FileText className="w-4 h-4 text-blue-500" />
                            )}
                          </div>
                          <div className="flex-1 min-w-0">
                            <h4 className="font-medium text-foreground text-sm truncate">{doc.filename}</h4>
                            <div className="flex items-center gap-2 text-xs text-muted-foreground">
                              <span>{getCategoryName(doc.document_type || '')}</span>
                              <span>•</span>
                              <span>{formatFileSize(doc.file_size)}</span>
                              <span>•</span>
                              <span>{new Date(doc.uploaded_at).toLocaleDateString()}</span>
                            </div>
                          </div>
                          {multiSelect && (
                            <Checkbox 
                              checked={isSelected(doc.id)}
                              className="flex-shrink-0"
                              onClick={(e) => e.stopPropagation()}
                              onCheckedChange={() => handleDocumentToggle(doc)}
                            />
                          )}
                        </div>
                      </CardContent>
                    </Card>
                  ))}
                </div>
              )}
            </ScrollArea>
          </div>
        </div>

        {/* Selected Documents Summary */}
        {selectedDocuments.length > 0 && (
          <div className="border-t pt-4 mt-4">
            <div className="flex items-center justify-between">
              <div className="text-sm">
                <span className="font-medium">{selectedDocuments.length}</span>
                <span className="text-muted-foreground"> document{selectedDocuments.length > 1 ? 's' : ''} selected</span>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                {selectedDocuments.slice(0, 3).map((doc) => (
                  <Badge key={doc.id} variant="secondary" className="text-xs">
                    {doc.filename.length > 20 ? doc.filename.slice(0, 20) + '...' : doc.filename}
                  </Badge>
                ))}
                {selectedDocuments.length > 3 && (
                  <Badge variant="secondary" className="text-xs">
                    +{selectedDocuments.length - 3} more
                  </Badge>
                )}
              </div>
            </div>
          </div>
        )}

        <DialogFooter>
          <Button variant="outline" onClick={handleClose}>
            Cancel
          </Button>
          <Button 
            onClick={handleConfirm} 
            disabled={selectedDocuments.length === 0}
          >
            <Check className="w-4 h-4 mr-2" />
            Attach {selectedDocuments.length > 0 ? `(${selectedDocuments.length})` : ''}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
