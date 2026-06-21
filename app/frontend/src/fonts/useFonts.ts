import { useEffect, useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import api from '../api/client';
import { registerFonts, type FontMeta } from './registry';

/**
 * Fetches the admin-uploaded font list and registers all fonts (uploaded +
 * built-in) with the browser and FortuneSheet's picker.
 *
 * `version` increments after every successful registration so the spreadsheet
 * can remount its canvas and redraw with the freshly loaded fonts.
 */
export function useFonts() {
  const [version, setVersion] = useState(0);

  const { data } = useQuery<FontMeta[]>({
    queryKey: ['fonts'],
    queryFn: () => api.get('/fonts').then((r) => r.data),
    staleTime: 5 * 60 * 1000,
  });

  useEffect(() => {
    let cancelled = false;
    registerFonts(data || []).finally(() => {
      if (!cancelled) setVersion((v) => v + 1);
    });
    return () => { cancelled = true; };
  }, [data]);

  return { fonts: data || [], version };
}
