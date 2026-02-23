/**
 * D365 AI Data Manager - SVG Icons
 * TOS-safe generic icons (NO Microsoft logos)
 * All 16x16, stroke=currentColor for theme compatibility
 */

const SVGIcons = {
  // 4-square grid for Power Apps
  powerApps: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" xmlns="http://www.w3.org/2000/svg">
    <rect x="1.5" y="1.5" width="5" height="5" rx="1"/>
    <rect x="9.5" y="1.5" width="5" height="5" rx="1"/>
    <rect x="1.5" y="9.5" width="5" height="5" rx="1"/>
    <rect x="9.5" y="9.5" width="5" height="5" rx="1"/>
  </svg>`,

  // Lightning bolt for Power Automate
  powerAutomate: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" xmlns="http://www.w3.org/2000/svg">
    <path d="M9 1L3 9h4l-1 6 6-8H8l1-6z"/>
  </svg>`,

  // Bar chart for Power BI
  powerBI: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" xmlns="http://www.w3.org/2000/svg">
    <rect x="1.5" y="9" width="3" height="6" rx="0.5"/>
    <rect x="6.5" y="5" width="3" height="10" rx="0.5"/>
    <rect x="11.5" y="1" width="3" height="14" rx="0.5"/>
  </svg>`,

  // Sparkle for AI Assistant
  aiAssistant: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" xmlns="http://www.w3.org/2000/svg">
    <path d="M8 1v3M8 12v3M1 8h3M12 8h3M3.5 3.5l2 2M10.5 10.5l2 2M12.5 3.5l-2 2M5.5 10.5l-2 2" stroke-linecap="round"/>
  </svg>`,

  // Trend line with dot for AI Analyze
  aiAnalyze: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" xmlns="http://www.w3.org/2000/svg">
    <path d="M1.5 13.5l4-5 3 2 6-8" stroke-linecap="round" stroke-linejoin="round"/>
    <circle cx="13" cy="3.5" r="1.5" fill="currentColor" stroke="none"/>
  </svg>`,

  // Generic connector icon for Power Platform button
  powerPlatform: `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" xmlns="http://www.w3.org/2000/svg">
    <circle cx="4" cy="4" r="2.5"/>
    <circle cx="12" cy="4" r="2.5"/>
    <circle cx="8" cy="12" r="2.5"/>
    <path d="M6 5l2 5M10 5l-2 5" stroke-linecap="round"/>
  </svg>`
};
