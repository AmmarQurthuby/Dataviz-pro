import { ChevronRight, Home } from "lucide-react";
import { cn } from "@/lib/utils";

interface BreadcrumbItem {
  label: string;
  href?: string;
  isActive?: boolean;
}

interface BreadcrumbProps {
  items: BreadcrumbItem[];
  className?: string;
}

export default function Breadcrumb({ items, className }: BreadcrumbProps) {
  return (
    <nav 
      aria-label="Breadcrumb"
      className={cn("flex items-center space-x-1 text-sm", className)}
    >
      <a 
        href="/" 
        className="flex items-center text-bps-gray-500 hover:text-bps-navy transition-colors"
        aria-label="Home"
      >
        <Home className="w-4 h-4" />
      </a>
      
      {items.map((item, index) => (
        <div key={index} className="flex items-center space-x-1">
          <ChevronRight className="w-4 h-4 text-bps-gray-400" />
          {item.href && !item.isActive ? (
            <a 
              href={item.href}
              className="text-bps-gray-500 hover:text-bps-navy transition-colors"
            >
              {item.label}
            </a>
          ) : (
            <span 
              className={cn(
                item.isActive 
                  ? "text-bps-navy font-medium" 
                  : "text-bps-gray-500"
              )}
            >
              {item.label}
            </span>
          )}
        </div>
      ))}
    </nav>
  );
}
