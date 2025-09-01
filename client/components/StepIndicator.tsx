import { Check } from "lucide-react";
import { cn } from "@/lib/utils";

interface StepIndicatorProps {
  currentStep: number;
  steps: string[];
}

export default function StepIndicator({ currentStep, steps }: StepIndicatorProps) {
  return (
    <div className="w-full py-6 bg-white border-b border-bps-gray-200">
      <div className="bps-container">
        <div className="flex items-center justify-between">
          {steps.map((step, index) => {
            const stepNumber = index + 1;
            const isCompleted = stepNumber < currentStep;
            const isActive = stepNumber === currentStep;
            const isPending = stepNumber > currentStep;

            return (
              <div key={index} className="flex items-center flex-1">
                {/* Step Circle */}
                <div className="flex items-center">
                  <div
                    className={cn(
                      "bps-step",
                      {
                        "completed": isCompleted,
                        "active": isActive,
                        "pending": isPending,
                      }
                    )}
                  >
                    {isCompleted ? (
                      <Check className="w-5 h-5" />
                    ) : (
                      <span>{stepNumber}</span>
                    )}
                  </div>

                  {/* Step Label */}
                  <div className="ml-3 hidden sm:block">
                    <p
                      className={cn(
                        "text-sm font-medium",
                        {
                          "text-bps-success": isCompleted,
                          "text-bps-navy": isActive,
                          "text-bps-gray-500": isPending,
                        }
                      )}
                    >
                      {step}
                    </p>
                    <p className="text-xs text-bps-gray-400">
                      Langkah {stepNumber}
                    </p>
                  </div>
                </div>

                {/* Connector Line */}
                {index < steps.length - 1 && (
                  <div className="flex-1 mx-4">
                    <div
                      className={cn(
                        "h-0.5 w-full",
                        isCompleted || (isActive && index < steps.length - 1)
                          ? "bg-bps-success"
                          : "bg-bps-gray-200"
                      )}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Mobile Step Labels */}
        <div className="sm:hidden mt-4">
          <p className="text-center text-sm font-medium text-bps-navy">
            {steps[currentStep - 1]} (Langkah {currentStep} dari {steps.length})
          </p>
        </div>
      </div>
    </div>
  );
}
