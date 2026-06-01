'use client';
import {
  Chart as ChartJS,
  CategoryScale, LinearScale, BarElement, ArcElement, PointElement, LineElement,
  Title, Tooltip, Legend, Filler,
} from 'chart.js';

ChartJS.register(
  CategoryScale, LinearScale, BarElement, ArcElement, PointElement, LineElement,
  Title, Tooltip, Legend, Filler,
);

ChartJS.defaults.font.family = "'Montserrat', sans-serif";
ChartJS.defaults.color = '#6b7280';
