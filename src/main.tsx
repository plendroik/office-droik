import { createRoot } from 'react-dom/client';
import { App } from './App';
import './styles.css';

// StrictMode kullanılmıyor: Univer ve pdf.js örneklerini çift başlatıp hemen yok etmek gereksiz yük getiriyor.
createRoot(document.getElementById('root')!).render(<App />);
