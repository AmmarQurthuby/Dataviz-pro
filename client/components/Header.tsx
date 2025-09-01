import React from "react";
import { Globe } from "lucide-react";

export default function Header() {
  return (
    <header className="bg-gradient-to-r from-bps-navy via-bps-navy-dark to-bps-navy text-white shadow-xl">
      {/* Top Bar with Language Selector */}
      <div className="border-b border-bps-navy-dark/30 bg-black/10">
        <div className="bps-container">
          <div className="flex items-center justify-between py-2 text-sm">
            <div className="flex items-center space-x-4">
              <span className="hidden md:inline font-medium">Badan Pusat Statistik Kota Medan</span>
              <span className="md:hidden font-medium">BPS</span>
            </div>
            <div className="flex items-center space-x-4">
              <button className="flex items-center space-x-2 px-3 py-1 rounded-full bg-white/10 hover:bg-white/20 transition-all duration-300 backdrop-blur-sm">
                <Globe className="w-4 h-4" />
                <span className="font-medium">ID</span>
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Main Header */}
      <div className="bps-container">
        <div className="flex items-center justify-between py-6">
          {/* Logo and Title */}
          <div className="flex items-center space-x-6">
            <div className="flex items-center space-x-4">
              {/* Logo Container */}
              <div className="relative">
                <div className="w-16 h-16 bg-white rounded-2xl flex items-center justify-center shadow-lg transform hover:scale-105 transition-transform duration-300">
                  <img 
                    src="bpslogo.svg" 
                    alt="BPS Logo" 
                    className="w-12 h-12 object-contain"
                  />
                </div>
                {/* Decorative ring */}
                <div className="absolute -inset-1 bg-gradient-to-r from-bps-blue-light to-transparent rounded-2xl blur opacity-30"></div>
              </div>
              
              <div>
                <h1 className="text-2xl md:text-3xl font-bold bg-gradient-to-r from-white to-bps-blue-light bg-clip-text text-transparent">
                  <span className="hidden md:inline">SISTEM VISUALISASI DATA</span>
                  <span className="md:hidden">VISUALISASI DATA</span>
                </h1>
                <div className="flex items-center space-x-2 mt-1">
                  <div className="h-0.5 w-8 bg-gradient-to-r from-bps-blue-light to-transparent"></div>
                  <p className="text-bps-blue-light text-sm font-medium hidden sm:block">
                    Portal Analisis Data Statistik
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* Decorative Element */}
          <div className="hidden lg:flex items-center space-x-4">
            <div className="flex space-x-2">
              <div className="w-2 h-2 bg-bps-blue-light rounded-full animate-pulse"></div>
              <div className="w-2 h-2 bg-white/60 rounded-full animate-pulse" style={{animationDelay: '0.5s'}}></div>
              <div className="w-2 h-2 bg-bps-blue-light/60 rounded-full animate-pulse" style={{animationDelay: '1s'}}></div>
            </div>
          </div>
        </div>
      </div>

      {/* Bottom decorative border */}
      <div className="h-1 bg-gradient-to-r from-transparent via-bps-blue-light to-transparent opacity-50"></div>
    </header>
  );
}