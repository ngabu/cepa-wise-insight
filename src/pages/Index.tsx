import { useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '@/contexts/AuthContext';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Leaf, Users, FileText, BarChart3 } from 'lucide-react';

const Index = () => {
  const { profile, loading } = useAuth();
  const navigate = useNavigate();

  useEffect(() => {
    if (loading) return;
    
    if (profile) {
      // Redirect authenticated users to their dashboard
      const { user_type, staff_unit, staff_position } = profile;
      
      if (['super_admin', 'admin'].includes(user_type)) {
        navigate('/admin', { replace: true });
      } else if (user_type === 'public') {
        navigate('/dashboard', { replace: true });
      } else if (user_type === 'staff' && staff_unit) {
        const isManagement = staff_position && ['manager', 'director', 'managing_director'].includes(staff_position);
        
        switch (staff_unit) {
          case 'compliance':
            navigate(isManagement ? '/ComplianceDashboard' : '/compliance', { replace: true });
            break;
          case 'registry':
            navigate(isManagement ? '/RegistryDashboard' : '/registry', { replace: true });
            break;
          case 'revenue':
            navigate(isManagement ? '/RevenueDashboard' : '/revenue', { replace: true });
            break;
          case 'finance':
            navigate(isManagement ? '/FinanceDashboard' : '/finance', { replace: true });
            break;
          case 'directorate':
            navigate('/directorate', { replace: true });
            break;
          default:
            navigate('/unauthorized', { replace: true });
        }
      } else {
        navigate('/unauthorized', { replace: true });
      }
    }
  }, [profile, loading, navigate]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-primary" />
      </div>
    );
  }

  // Show landing page for non-authenticated users
  if (!profile) {
    return (
      <div className="min-h-screen bg-background">
        {/* Header */}
        <header className="bg-card/80 backdrop-blur-sm border-b border-border">
          <div className="container mx-auto px-4 py-4 flex justify-between items-center">
            <div className="flex items-center gap-2">
              <div className="p-2 bg-gradient-primary rounded-lg shadow-primary">
                <Leaf className="w-6 h-6 text-primary-foreground" />
              </div>
              <h1 className="text-xl font-bold text-foreground">PNG Conservation And Environment Protection Authority</h1>
            </div>
            <Button 
              onClick={() => navigate('/auth')}
              variant="secondary"
            >
              Sign In
            </Button>
          </div>
        </header>

        {/* Hero Section */}
        <main className="container mx-auto px-4 py-16">
          <div className="text-center mb-16">
            <h2 className="text-4xl md:text-6xl font-bold text-foreground mb-6">
              Environmental Management
              <span className="block text-primary">Made Simple</span>
            </h2>
            <p className="text-xl text-muted-foreground mb-8 max-w-2xl mx-auto">
              Streamline permit applications, manage compliance, and protect Papua New Guinea's natural resources with our comprehensive digital platform.
            </p>
            <Button 
              size="lg" 
              onClick={() => navigate('/auth?tab=signup')}
              variant="gradient"
              className="text-lg px-8 py-6"
            >
              Get Started
            </Button>
          </div>

          {/* Features Grid */}
          <div className="grid md:grid-cols-3 gap-8 mb-16">
            <Card className="hover:shadow-card transition-shadow">
              <CardHeader>
                <div className="w-12 h-12 bg-gradient-primary rounded-lg flex items-center justify-center mb-4 shadow-primary">
                  <FileText className="w-6 h-6 text-primary-foreground" />
                </div>
                <CardTitle>Digital Applications</CardTitle>
              </CardHeader>
              <CardContent>
                <CardDescription>
                  Submit and track permit applications online with our streamlined digital process.
                </CardDescription>
              </CardContent>
            </Card>

            <Card className="hover:shadow-card transition-shadow">
              <CardHeader>
                <div className="w-12 h-12 bg-gradient-primary rounded-lg flex items-center justify-center mb-4 shadow-primary">
                  <Users className="w-6 h-6 text-primary-foreground" />
                </div>
                <CardTitle>Multi-User Access</CardTitle>
              </CardHeader>
              <CardContent>
                <CardDescription>
                  Role-based access for public users, staff, and administrators with secure authentication.
                </CardDescription>
              </CardContent>
            </Card>

            <Card className="hover:shadow-card transition-shadow">
              <CardHeader>
                <div className="w-12 h-12 bg-gradient-primary rounded-lg flex items-center justify-center mb-4 shadow-primary">
                  <BarChart3 className="w-6 h-6 text-primary-foreground" />
                </div>
                <CardTitle>Real-time Monitoring</CardTitle>
              </CardHeader>
              <CardContent>
                <CardDescription>
                  Track compliance, monitor progress, and generate reports with comprehensive analytics.
                </CardDescription>
              </CardContent>
            </Card>
          </div>

          {/* About Section */}
          <div className="text-center max-w-4xl mx-auto">
            <h3 className="text-3xl font-bold text-foreground mb-6">
              About the Conservation & Environment Protection Authority
            </h3>
            <div className="space-y-4 text-muted-foreground text-lg leading-relaxed">
              <p>
                The PNG Conservation and Environment Protection Authority is the primary regulatory body responsible for environmental management and protection across Papua New Guinea. Our digital platform facilitates the entire lifecycle of environmental permit applications, assessments, and compliance monitoring.
              </p>
              <p>
                Through this system, businesses and individuals can submit permit applications for various environmental activities, track their application status, manage compliance requirements, and communicate with regulatory officers. Our multi-unit approach ensures specialized handling across Registry, Compliance, Revenue, and Finance departments for efficient processing and oversight.
              </p>
              <p>
                The platform supports transparent governance, statutory deadline tracking, comprehensive reporting, and real-time monitoring to safeguard Papua New Guinea&apos;s rich natural resources while enabling sustainable development.
              </p>
            </div>
          </div>
        </main>
      </div>
    );
  }

  return null;
};

export default Index;
