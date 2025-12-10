import React, { useEffect, useRef, useState, useCallback } from 'react';
import mapboxgl from 'mapbox-gl';
import 'mapbox-gl/dist/mapbox-gl.css';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { MapPin, AlertTriangle, Clock } from 'lucide-react';
import { Switch } from '@/components/ui/switch';
import { Label } from '@/components/ui/label';
import { supabase } from '@/integrations/supabase/client';
import { Badge } from '@/components/ui/badge';

interface PermitWithFee {
  id: string;
  title: string | null;
  project_title: string | null;
  entity_name: string | null;
  latitude: number;
  longitude: number;
  due_date: string;
  amount: number;
  invoice_id: string;
  payment_status: string;
  is_overdue: boolean;
  is_due_soon: boolean;
}

export function RevenueFeeMap() {
  const mapContainer = useRef<HTMLDivElement>(null);
  const map = useRef<mapboxgl.Map | null>(null);
  const markers = useRef<mapboxgl.Marker[]>([]);
  const popupRef = useRef<mapboxgl.Popup | null>(null);
  
  const [mapLoaded, setMapLoaded] = useState(false);
  const [permits, setPermits] = useState<PermitWithFee[]>([]);
  const [loading, setLoading] = useState(true);
  
  // Toggles
  const [showOverdue, setShowOverdue] = useState(true);
  const [showDueSoon, setShowDueSoon] = useState(true);

  // Fetch permits with fee data
  const fetchPermitsWithFees = useCallback(async () => {
    setLoading(true);
    try {
      const now = new Date();
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);

      // Fetch invoices with related permit and intent registration location data
      const { data: invoices, error: invoicesError } = await supabase
        .from('invoices')
        .select(`
          id,
          amount,
          due_date,
          payment_status,
          invoice_type,
          permit_id,
          intent_registration_id,
          entity_id,
          entities:entity_id (name),
          permit_applications:permit_id (
            id,
            title,
            coordinates,
            intent_registration_id
          )
        `)
        .in('payment_status', ['pending', 'overdue'])
        .order('due_date', { ascending: true });

      if (invoicesError) {
        console.error('Error fetching invoices:', invoicesError);
        return;
      }

      const permitsWithLocation: PermitWithFee[] = [];

      for (const invoice of invoices || []) {
        let lat: number | null = null;
        let lng: number | null = null;
        let projectTitle: string | null = null;

        // Try to get coordinates from permit_applications first
        const permit = invoice.permit_applications as any;
        if (permit?.coordinates) {
          const coords = typeof permit.coordinates === 'string' 
            ? JSON.parse(permit.coordinates) 
            : permit.coordinates;
          lat = coords.lat || coords.latitude;
          lng = coords.lng || coords.longitude;
          projectTitle = permit.title;
        }

        // If no coordinates from permit, try to get from linked intent_registration
        if (!lat || !lng) {
          const intentId = permit?.intent_registration_id || invoice.intent_registration_id;
          if (intentId) {
            const { data: intent } = await supabase
              .from('intent_registrations')
              .select('project_title, latitude, longitude')
              .eq('id', intentId)
              .single();
            
            if (intent?.latitude && intent?.longitude) {
              lat = intent.latitude;
              lng = intent.longitude;
              projectTitle = intent.project_title;
            }
          }
        }

        // Only include permits with valid coordinates
        if (lat && lng) {
          const dueDate = new Date(invoice.due_date);
          const isOverdue = dueDate < now && invoice.payment_status !== 'paid';
          const isDueSoon = dueDate >= now && dueDate <= thirtyDaysFromNow && invoice.payment_status !== 'paid';

          // Only include if it matches our filter criteria
          if (isOverdue || isDueSoon) {
            const entity = invoice.entities as any;
            permitsWithLocation.push({
              id: permit?.id || invoice.id,
              title: permit?.title || null,
              project_title: projectTitle,
              entity_name: entity?.name || null,
              latitude: lat,
              longitude: lng,
              due_date: invoice.due_date,
              amount: invoice.amount,
              invoice_id: invoice.id,
              payment_status: invoice.payment_status,
              is_overdue: isOverdue,
              is_due_soon: isDueSoon
            });
          }
        }
      }

      // If no invoices have location data, generate sample data for demonstration
      if (permitsWithLocation.length === 0) {
        // Use default PNG locations for demo
        const sampleLocations = [
          { lat: -6.314993, lng: 147.1494, title: 'Lae Industrial Permit', entity: 'Pacific Industries Ltd', overdue: true },
          { lat: -5.48765, lng: 143.71905, title: 'Enga Mining Permit', entity: 'Highland Mining Co.', overdue: true },
          { lat: -9.4324, lng: 147.18, title: 'Port Moresby Waste Permit', entity: 'Capital Waste Management', overdue: false },
          { lat: -6.73075, lng: 147.016, title: 'Morobe Forestry Permit', entity: 'Morobe Timber Ltd', overdue: false },
          { lat: -5.22, lng: 145.78, title: 'Madang Fishing Permit', entity: 'Madang Fisheries', overdue: true },
        ];

        sampleLocations.forEach((loc, index) => {
          const daysOffset = loc.overdue ? -(Math.floor(Math.random() * 30) + 1) : Math.floor(Math.random() * 25) + 5;
          const dueDate = new Date();
          dueDate.setDate(dueDate.getDate() + daysOffset);

          permitsWithLocation.push({
            id: `sample-${index}`,
            title: loc.title,
            project_title: loc.title,
            entity_name: loc.entity,
            latitude: loc.lat,
            longitude: loc.lng,
            due_date: dueDate.toISOString(),
            amount: Math.floor(Math.random() * 100000) + 10000,
            invoice_id: `inv-sample-${index}`,
            payment_status: loc.overdue ? 'overdue' : 'pending',
            is_overdue: loc.overdue,
            is_due_soon: !loc.overdue
          });
        });
      }

      setPermits(permitsWithLocation);
    } catch (error) {
      console.error('Error fetching permits with fees:', error);
    } finally {
      setLoading(false);
    }
  }, []);

  // Initialize map
  useEffect(() => {
    if (!mapContainer.current || map.current) return;

    mapboxgl.accessToken = 'pk.eyJ1IjoibG92YWJsZS1zdGFuZG9uZSIsImEiOiJjbHVwcmh3N3MwdGNtMnByeHg5MTdzeXdzIn0.d3dOj93xnwgDqHiSF2W1Dw';
    
    map.current = new mapboxgl.Map({
      container: mapContainer.current,
      style: 'mapbox://styles/mapbox/light-v11',
      center: [147.1494, -6.314993], // PNG center
      zoom: 5.5
    });

    map.current.addControl(new mapboxgl.NavigationControl(), 'top-right');

    map.current.on('load', () => {
      setMapLoaded(true);
    });

    // Remove popup on map click (empty area)
    map.current.on('click', (e) => {
      // Check if click is on the map background (not on a marker)
      if (popupRef.current) {
        popupRef.current.remove();
        popupRef.current = null;
      }
    });

    return () => {
      if (map.current) {
        map.current.remove();
        map.current = null;
      }
    };
  }, []);

  // Fetch data on mount
  useEffect(() => {
    fetchPermitsWithFees();
  }, [fetchPermitsWithFees]);

  // Update markers when permits or toggles change
  useEffect(() => {
    if (!map.current || !mapLoaded) return;

    // Clear existing markers
    markers.current.forEach(marker => marker.remove());
    markers.current = [];

    // Remove any existing popup
    if (popupRef.current) {
      popupRef.current.remove();
      popupRef.current = null;
    }

    // Filter permits based on toggles
    const filteredPermits = permits.filter(permit => {
      if (permit.is_overdue && !showOverdue) return false;
      if (permit.is_due_soon && !showDueSoon) return false;
      return true;
    });

    // Add markers for each permit
    filteredPermits.forEach(permit => {
      // Create custom marker element
      const el = document.createElement('div');
      el.className = 'revenue-fee-marker';
      el.style.width = '24px';
      el.style.height = '24px';
      el.style.borderRadius = '50%';
      el.style.border = '3px solid white';
      el.style.boxShadow = '0 2px 6px rgba(0,0,0,0.3)';
      el.style.cursor = 'pointer';
      
      // Red for overdue, amber for due soon
      if (permit.is_overdue) {
        el.style.backgroundColor = '#dc2626'; // red-600
      } else {
        el.style.backgroundColor = '#f59e0b'; // amber-500
      }

      const marker = new mapboxgl.Marker({ element: el })
        .setLngLat([permit.longitude, permit.latitude])
        .addTo(map.current!);

      // Show popup on hover
      el.addEventListener('mouseenter', () => {
        // Remove any existing popup first
        if (popupRef.current) {
          popupRef.current.remove();
        }

        const dueDate = new Date(permit.due_date);
        const formattedDate = dueDate.toLocaleDateString('en-AU', { 
          day: 'numeric', 
          month: 'short', 
          year: 'numeric' 
        });
        const formattedAmount = new Intl.NumberFormat('en-PG', {
          style: 'currency',
          currency: 'PGK',
          minimumFractionDigits: 0
        }).format(permit.amount);

        const statusLabel = permit.is_overdue 
          ? '<span style="color: #dc2626; font-weight: 600;">OVERDUE</span>'
          : '<span style="color: #f59e0b; font-weight: 600;">Due Soon</span>';

        const popupContent = `
          <div style="padding: 8px; min-width: 200px;">
            <div style="font-weight: 600; font-size: 14px; margin-bottom: 6px; color: #1f2937;">
              ${permit.project_title || permit.title || 'Permit Fee'}
            </div>
            <div style="font-size: 12px; color: #6b7280; margin-bottom: 4px;">
              <strong>Entity:</strong> ${permit.entity_name || 'N/A'}
            </div>
            <div style="font-size: 12px; color: #6b7280; margin-bottom: 4px;">
              <strong>Amount:</strong> ${formattedAmount}
            </div>
            <div style="font-size: 12px; color: #6b7280; margin-bottom: 4px;">
              <strong>Due:</strong> ${formattedDate}
            </div>
            <div style="font-size: 12px; margin-top: 6px;">
              ${statusLabel}
            </div>
          </div>
        `;

        popupRef.current = new mapboxgl.Popup({
          closeButton: false,
          closeOnClick: false,
          offset: 15
        })
          .setLngLat([permit.longitude, permit.latitude])
          .setHTML(popupContent)
          .addTo(map.current!);
      });

      // Hide popup on mouse leave
      el.addEventListener('mouseleave', () => {
        if (popupRef.current) {
          popupRef.current.remove();
          popupRef.current = null;
        }
      });

      markers.current.push(marker);
    });

    // Fit bounds to show all markers if there are any
    if (filteredPermits.length > 0) {
      const bounds = new mapboxgl.LngLatBounds();
      filteredPermits.forEach(permit => {
        bounds.extend([permit.longitude, permit.latitude]);
      });
      map.current.fitBounds(bounds, { padding: 50, maxZoom: 8 });
    }
  }, [permits, mapLoaded, showOverdue, showDueSoon]);

  const overdueCount = permits.filter(p => p.is_overdue).length;
  const dueSoonCount = permits.filter(p => p.is_due_soon).length;

  return (
    <Card>
      <CardHeader className="pb-3">
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div>
            <CardTitle className="flex items-center gap-2 text-lg">
              <MapPin className="h-5 w-5 text-primary" />
              Annual Permit Fee Locations
            </CardTitle>
            <CardDescription className="mt-1">
              Geographical view of permits with overdue or upcoming annual fees
            </CardDescription>
          </div>
          <div className="flex flex-wrap gap-3">
            <Badge variant="destructive" className="flex items-center gap-1">
              <AlertTriangle className="h-3 w-3" />
              {overdueCount} Overdue
            </Badge>
            <Badge variant="outline" className="flex items-center gap-1 border-amber-500 text-amber-600">
              <Clock className="h-3 w-3" />
              {dueSoonCount} Due in 30 days
            </Badge>
          </div>
        </div>
      </CardHeader>
      <CardContent className="space-y-4">
        {/* Toggle controls */}
        <div className="flex flex-wrap gap-4 p-3 bg-muted/50 rounded-lg">
          <div className="flex items-center space-x-2">
            <Switch
              id="show-overdue"
              checked={showOverdue}
              onCheckedChange={setShowOverdue}
            />
            <Label htmlFor="show-overdue" className="flex items-center gap-2 cursor-pointer">
              <div className="w-3 h-3 rounded-full bg-red-600" />
              <span className="text-sm">Overdue Fees</span>
            </Label>
          </div>
          <div className="flex items-center space-x-2">
            <Switch
              id="show-due-soon"
              checked={showDueSoon}
              onCheckedChange={setShowDueSoon}
            />
            <Label htmlFor="show-due-soon" className="flex items-center gap-2 cursor-pointer">
              <div className="w-3 h-3 rounded-full bg-amber-500" />
              <span className="text-sm">Due in 30 Days</span>
            </Label>
          </div>
        </div>

        {/* Map container */}
        <div className="relative">
          <div 
            ref={mapContainer} 
            className="w-full h-[400px] rounded-lg overflow-hidden border"
          />
          {loading && (
            <div className="absolute inset-0 flex items-center justify-center bg-background/80">
              <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary"></div>
            </div>
          )}
        </div>

        {/* Legend */}
        <div className="flex flex-wrap gap-4 text-xs text-muted-foreground">
          <div className="flex items-center gap-1">
            <div className="w-3 h-3 rounded-full bg-red-600 border-2 border-white shadow" />
            <span>Overdue annual permit fees</span>
          </div>
          <div className="flex items-center gap-1">
            <div className="w-3 h-3 rounded-full bg-amber-500 border-2 border-white shadow" />
            <span>Annual fees due within 30 days</span>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
