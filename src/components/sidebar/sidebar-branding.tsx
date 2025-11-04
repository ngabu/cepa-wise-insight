
import { TreePine } from "lucide-react"

interface SidebarBrandingProps {
  collapsed: boolean
  unitName: string
}

export function SidebarBranding({ collapsed, unitName }: SidebarBrandingProps) {
  return (
    <div className="">
      {!collapsed ? (
        <div className="flex items-center space-x-3">
          <div className="w-12 h-12 bg-white/10 rounded-2xl flex items-center justify-center backdrop-blur-sm border border-white/20">
            <TreePine className="w-7 h-7 text-white" />
          </div>
          <div>
            <h2 className="font-bold text-white text-xl tracking-tight">PNG Conservation</h2>
            <p className="text-sm text-white/80 font-medium mt-0.5">
              {unitName}
            </p>
          </div>
        </div>
      ) : (
        <div className="w-12 h-12 bg-white/10 rounded-2xl flex items-center justify-center mx-auto backdrop-blur-sm border border-white/20">
          <TreePine className="w-7 h-7 text-white" />
        </div>
      )}
    </div>
  )
}
