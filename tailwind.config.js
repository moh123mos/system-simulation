/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{vue,js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        primary: '#1E3A8A', // Custom blue
        secondary: '#9333EA', // Custom purple
        accent: '#FACC15', // Custom yellow
        customGray: {
          50: '#F9FAFB',
          100: '#F3F4F6',
          200: '#E5E7EB',
          900: '#111827', // Dark gray
        },
      },
      backgroundColor: {
        'custom-dark': '#121212',
        'custom-light': '#F4F4F4',
      },
    },
  },
  plugins: [],
}

