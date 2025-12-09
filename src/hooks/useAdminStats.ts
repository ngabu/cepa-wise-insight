import { useState, useEffect, useMemo } from 'react';
import { supabase } from '@/integrations/supabase/client';
import { useDateFilter, type DateFilterPeriod, type DateFilterRange } from './useDateFilter';

interface AdminStats {
  totalUsers: number;
  activeUsers: number;
  suspendedUsers: number;
  publicUsers: number;
  staffUsers: number;
  adminUsers: number;
  superAdminUsers: number;
  totalApplications: number;
  draftApplications: number;
  submittedApplications: number;
  underReviewApplications: number;
  approvedApplications: number;
  rejectedApplications: number;
  totalTransactions: number;
  pendingTransactions: number;
  completedTransactions: number;
  totalEntities: number;
}

export function useAdminStats(period?: DateFilterPeriod) {
  const dateRange = useDateFilter(period || 'all-time');
  
  const [stats, setStats] = useState<AdminStats>({
    totalUsers: 0,
    activeUsers: 0,
    suspendedUsers: 0,
    publicUsers: 0,
    staffUsers: 0,
    adminUsers: 0,
    superAdminUsers: 0,
    totalApplications: 0,
    draftApplications: 0,
    submittedApplications: 0,
    underReviewApplications: 0,
    approvedApplications: 0,
    rejectedApplications: 0,
    totalTransactions: 0,
    pendingTransactions: 0,
    completedTransactions: 0,
    totalEntities: 0,
  });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchStats();
  }, [period, dateRange.start.toISOString(), dateRange.end.toISOString()]);

  const fetchStats = async () => {
    try {
      setLoading(true);
      setError(null);

      // Fetch user statistics (users don't typically have created_at filtering for total counts)
      const { data: userStats, error: userError } = await supabase
        .from('profiles')
        .select('user_type, is_active');

      if (userError) throw userError;

      // Fetch application statistics with date filter
      let appQuery = supabase
        .from('permit_applications')
        .select('status, created_at');
      
      if (period && period !== 'all-time') {
        appQuery = appQuery
          .gte('created_at', dateRange.start.toISOString())
          .lte('created_at', dateRange.end.toISOString());
      }
      
      const { data: appStats, error: appError } = await appQuery;
      if (appError) throw appError;

      // Fetch financial transaction statistics with date filter
      let transactionQuery = supabase
        .from('financial_transactions')
        .select('status, amount, created_at');
      
      if (period && period !== 'all-time') {
        transactionQuery = transactionQuery
          .gte('created_at', dateRange.start.toISOString())
          .lte('created_at', dateRange.end.toISOString());
      }
      
      const { data: transactionStats, error: transactionError } = await transactionQuery;
      if (transactionError) throw transactionError;

      // Fetch entities statistics with date filter
      let entitiesQuery = supabase
        .from('entities')
        .select('id, created_at');
      
      if (period && period !== 'all-time') {
        entitiesQuery = entitiesQuery
          .gte('created_at', dateRange.start.toISOString())
          .lte('created_at', dateRange.end.toISOString());
      }
      
      const { data: entitiesStats, error: entitiesError } = await entitiesQuery;
      if (entitiesError) throw entitiesError;

      // Calculate user statistics (not date filtered - showing all-time user stats)
      const totalUsers = userStats?.length || 0;
      const activeUsers = userStats?.filter(u => u.is_active).length || 0;
      const suspendedUsers = userStats?.filter(u => !u.is_active).length || 0;
      const publicUsers = userStats?.filter(u => u.user_type === 'public').length || 0;
      const staffUsers = userStats?.filter(u => u.user_type === 'staff').length || 0;
      const adminUsers = userStats?.filter(u => u.user_type === 'admin').length || 0;
      const superAdminUsers = userStats?.filter(u => u.user_type === 'super_admin').length || 0;

      // Calculate application statistics (date filtered)
      const totalApplications = appStats?.length || 0;
      const draftApplications = appStats?.filter(a => a.status === 'draft').length || 0;
      const submittedApplications = appStats?.filter(a => a.status === 'submitted').length || 0;
      const underReviewApplications = appStats?.filter(a => a.status === 'under_review').length || 0;
      const approvedApplications = appStats?.filter(a => a.status === 'approved').length || 0;
      const rejectedApplications = appStats?.filter(a => a.status === 'rejected').length || 0;

      // Calculate transaction statistics (date filtered)
      const totalTransactions = transactionStats?.length || 0;
      const pendingTransactions = transactionStats?.filter(t => t.status === 'pending').length || 0;
      const completedTransactions = transactionStats?.filter(t => t.status === 'completed').length || 0;
      
      // Calculate entities statistics (date filtered)
      const totalEntities = entitiesStats?.length || 0;

      setStats({
        totalUsers,
        activeUsers,
        suspendedUsers,
        publicUsers,
        staffUsers,
        adminUsers,
        superAdminUsers,
        totalApplications,
        draftApplications,
        submittedApplications,
        underReviewApplications,
        approvedApplications,
        rejectedApplications,
        totalTransactions,
        pendingTransactions,
        completedTransactions,
        totalEntities,
      });
    } catch (err) {
      console.error('Error fetching admin stats:', err);
      setError(err instanceof Error ? err.message : 'Failed to fetch statistics');
    } finally {
      setLoading(false);
    }
  };

  return { stats, loading, error, refetch: fetchStats };
}