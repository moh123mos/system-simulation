import { watchEffect, onMounted, onUnmounted } from 'vue';
import { useTheme } from 'vuetify';

export function useBodyBackground() {
  const theme = useTheme();

  const updateBackgroundColor = () => {
    const backgroundColor = theme.global.current.value.colors.background;
    document.body.style.backgroundColor = backgroundColor;
  };

  // Watch for theme changes and update the body background color
  watchEffect(updateBackgroundColor);

  onMounted(updateBackgroundColor);
  onUnmounted(() => {
    document.body.style.backgroundColor = ''; // Reset on unmount
  });
}
