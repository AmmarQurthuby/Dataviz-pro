import React from "react";
import { Mail, Phone, MapPin } from "lucide-react";

export default function Footer() {
  return (
    <footer className="bg-gradient-to-b from-bps-navy via-bps-navy-dark to-gray-900 text-white">
      {/* Top decorative border */}
      <div className="h-1 bg-gradient-to-r from-transparent via-bps-blue-light to-transparent opacity-50"></div>
      
      {/* Main Footer Content */}
      <div className="bps-container py-12">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          
          {/* Logo and Description */}
          <div className="lg:col-span-2">
            <div className="flex items-center space-x-4 mb-6">
              <div className="w-16 h-16 bg-white rounded-2xl flex items-center justify-center shadow-lg">
                <img 
                  src="/bpslogo.svg" 
                  alt="BPS Logo" 
                  className="w-12 h-12 object-contain"
                />
              </div>
              <div>
                <h3 className="text-xl font-bold bg-gradient-to-r from-white to-bps-blue-light bg-clip-text text-transparent">
                  SISTEM VISUALISASI DATA
                </h3>
                <p className="text-bps-blue-light text-sm">
                  Badan Pusat Statistik Kota Medan
                </p>
              </div>
            </div>
            <p className="text-gray-300 leading-relaxed mb-6 max-w-md">
              Portal analisis data statistik yang menyediakan visualisasi data interaktif 
              untuk mendukung pengambilan keputusan berbasis data statistik Indonesia.
            </p>
          </div>

          {/* Contact Info */}
          <div>
            <h4 className="text-lg font-semibold mb-6 text-white">
              Kontak
            </h4>
            <div className="space-y-4">
              <div className="mb-4">
                <p className="text-gray-300 text-sm font-medium leading-relaxed">
                  Badan Pusat Statistik Kota Medan<br />
                  <span className="text-gray-400 text-xs">(Statistics of Medan Municipality)</span>
                </p>
              </div>
              <div className="flex items-start space-x-3">
                <MapPin className="w-5 h-5 text-bps-blue-light mt-0.5 flex-shrink-0" />
                <div>
                  <p className="text-gray-300 text-sm leading-relaxed">
                    Jl. Gaperta/ Brigjend. H. Abdul Manaf Lubis No. 311<br />
                    Medan - Indonesia
                  </p>
                </div>
              </div>
              <div className="flex items-center space-x-3">
                <Phone className="w-5 h-5 text-bps-blue-light flex-shrink-0" />
                <div>
                  <p className="text-gray-300 text-sm">(62-61) 8449289</p>
                </div>
              </div>
              <div className="flex items-center space-x-3">
                <Mail className="w-5 h-5 text-bps-blue-light flex-shrink-0" />
                <div>
                  <a href="mailto:bps1275@bps.go.id" className="text-gray-300 hover:text-bps-blue-light transition-colors text-sm">
                    bps1275@bps.go.id
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Bottom Bar */}
      <div className="border-t border-gray-700/50 bg-black/20">
        <div className="bps-container">
          <div className="flex flex-col md:flex-row items-center justify-between py-6 space-y-4 md:space-y-0">
            <div className="flex items-center space-x-4 text-sm text-gray-400">
              <span>&copy; {new Date().getFullYear()} Badan Pusat Statistik Kota Medan</span>
            </div>
            
            <div className="flex items-center space-x-2">
              <div className="w-2 h-2 bg-green-400 rounded-full animate-pulse"></div>
              <span className="text-green-400 text-xs font-medium">Sistem Aktif</span>
            </div>
          </div>
        </div>
      </div>

      {/* Background Pattern */}
      <div className="absolute inset-0 opacity-5 pointer-events-none">
        <div className="absolute inset-0" style={{
          backgroundImage: `radial-gradient(circle at 25% 25%, rgba(59, 130, 246, 0.1) 0%, transparent 50%),
                           radial-gradient(circle at 75% 75%, rgba(59, 130, 246, 0.1) 0%, transparent 50%)`
        }}></div>
      </div>
    </footer>
  );
}