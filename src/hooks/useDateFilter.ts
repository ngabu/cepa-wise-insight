import { useMemo } from 'react';
import { startOfMonth, startOfYear, subDays, subMonths, subYears, endOfYear, startOfQuarter, subQuarters } from 'date-fns';

export type DateFilterPeriod = 'weekly' | 'monthly' | 'quarterly' | 'yearly' | 'mtd' | 'ytd' | 'last-year' | 'all-time';

// Helper to get trend data labels based on period
export function getTrendLabelsForPeriod(period: DateFilterPeriod, dateRange: DateFilterRange): string[] {
  const { start, end } = dateRange;
  const labels: string[] = [];
  
  switch (period) {
    case 'weekly': {
      // Show daily labels for 7 days
      const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      for (let i = 6; i >= 0; i--) {
        const date = subDays(end, i);
        labels.push(days[date.getDay()]);
      }
      return labels;
    }
    case 'monthly': {
      // Show last 4 weeks or daily for 30 days
      return ['Week 1', 'Week 2', 'Week 3', 'Week 4'];
    }
    case 'quarterly': {
      // Show 3 months
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      for (let i = 2; i >= 0; i--) {
        const date = subMonths(end, i);
        labels.push(months[date.getMonth()]);
      }
      return labels;
    }
    case 'yearly':
    case 'ytd':
    case 'last-year': {
      // Show 12 months
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      if (period === 'last-year') {
        return months;
      }
      // For ytd, show months from start of year to now
      const startMonth = start.getMonth();
      const endMonth = end.getMonth();
      for (let i = startMonth; i <= endMonth; i++) {
        labels.push(months[i]);
      }
      return labels.length > 0 ? labels : months;
    }
    case 'mtd': {
      // Show weeks of current month
      return ['Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5'];
    }
    case 'all-time': {
      // Show years
      const startYear = start.getFullYear();
      const endYear = end.getFullYear();
      for (let y = startYear; y <= endYear; y++) {
        labels.push(y.toString());
      }
      return labels;
    }
    default:
      return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  }
}

// Helper to check if a date falls within a trend bucket
export function getDataBucketIndex(
  itemDate: Date,
  period: DateFilterPeriod,
  dateRange: DateFilterRange
): number {
  const { start, end } = dateRange;
  
  switch (period) {
    case 'weekly': {
      const daysDiff = Math.floor((end.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24));
      return daysDiff >= 0 && daysDiff < 7 ? 6 - daysDiff : -1;
    }
    case 'monthly': {
      const daysDiff = Math.floor((end.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24));
      if (daysDiff < 0 || daysDiff >= 30) return -1;
      return Math.floor((29 - daysDiff) / 7);
    }
    case 'quarterly': {
      const monthDiff = (end.getFullYear() - itemDate.getFullYear()) * 12 + (end.getMonth() - itemDate.getMonth());
      return monthDiff >= 0 && monthDiff < 3 ? 2 - monthDiff : -1;
    }
    case 'yearly':
    case 'ytd': {
      if (itemDate < start || itemDate > end) return -1;
      return itemDate.getMonth() - start.getMonth();
    }
    case 'last-year': {
      if (itemDate < start || itemDate > end) return -1;
      return itemDate.getMonth();
    }
    case 'mtd': {
      if (itemDate < start || itemDate > end) return -1;
      const weekOfMonth = Math.floor((itemDate.getDate() - 1) / 7);
      return Math.min(weekOfMonth, 4);
    }
    case 'all-time': {
      if (itemDate < start || itemDate > end) return -1;
      return itemDate.getFullYear() - start.getFullYear();
    }
    default:
      return -1;
  }
}

export interface DateFilterRange {
  start: Date;
  end: Date;
}

export function useDateFilter(period: DateFilterPeriod): DateFilterRange {
  return useMemo(() => {
    const now = new Date();
    
    switch (period) {
      case 'weekly':
        return { start: subDays(now, 7), end: now };
      case 'monthly':
        return { start: subDays(now, 30), end: now };
      case 'quarterly':
        return { start: subMonths(now, 3), end: now };
      case 'yearly':
        return { start: subYears(now, 1), end: now };
      case 'mtd':
        return { start: startOfMonth(now), end: now };
      case 'ytd':
        return { start: startOfYear(now), end: now };
      case 'last-year':
        return { start: startOfYear(subYears(now, 1)), end: endOfYear(subYears(now, 1)) };
      case 'all-time':
        return { start: new Date(2020, 0, 1), end: now };
      default:
        return { start: subDays(now, 30), end: now };
    }
  }, [period]);
}

export function filterByDateRange<T extends { created_at?: string | null }>(
  data: T[] | null | undefined,
  dateRange: DateFilterRange,
  dateField: keyof T = 'created_at' as keyof T
): T[] {
  if (!data) return [];
  
  return data.filter(item => {
    const dateValue = item[dateField];
    if (!dateValue || typeof dateValue !== 'string') return false;
    
    const itemDate = new Date(dateValue);
    return itemDate >= dateRange.start && itemDate <= dateRange.end;
  });
}
