import { useState } from "react";
import { ChevronDown, ChevronUp } from "lucide-react";
import { cn } from "@/lib/utils";

interface CollapsibleSectionProps {
  title: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
  className?: string;
  titleClassName?: string;
}

export default function CollapsibleSection({ 
  title, 
  children, 
  defaultOpen = true,
  className,
  titleClassName 
}: CollapsibleSectionProps) {
  const [isOpen, setIsOpen] = useState(defaultOpen);

  return (
    <div className={cn("bps-card", className)}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className={cn(
          "w-full flex items-center justify-between p-4 text-left font-semibold text-bps-navy hover:bg-bps-gray-50 transition-colors rounded-t-lg",
          !isOpen && "rounded-b-lg",
          titleClassName
        )}
      >
        <span>{title}</span>
        {isOpen ? (
          <ChevronUp className="w-5 h-5 text-bps-gray-500" />
        ) : (
          <ChevronDown className="w-5 h-5 text-bps-gray-500" />
        )}
      </button>
      
      {isOpen && (
        <div className="p-4 pt-0 animate-fade-in">
          {children}
        </div>
      )}
    </div>
  );
}
