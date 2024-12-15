function toggleDarkMode(isDarkMode) {
  isDarkMode.status = !isDarkMode.status;

  if (isDarkMode.status) {
    document.documentElement.classList.add('dark');
    localStorage.setItem('theme', 'dark'); // Save preference
  } else {
    document.documentElement.classList.remove('dark');
    localStorage.setItem('theme', 'light'); // Save preference
  }
}
export default toggleDarkMode;
