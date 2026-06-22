import { useEffect, useRef } from 'react';
import { io, Socket } from 'socket.io-client';
import { useAuth } from '../store/auth';

let colorSocket: Socket | null = null;

export function useSocketColorSync(onColorChange: (username: string, color: string) => void) {
  const { token } = useAuth();
  const cbRef = useRef(onColorChange);
  cbRef.current = onColorChange;

  useEffect(() => {
    if (!token) return;
    if (!colorSocket || colorSocket.disconnected) {
      colorSocket = io(import.meta.env.VITE_API_URL || window.location.origin, {
        auth: { token },
        path: '/socket.io',
        transports: ['websocket'],
      });
    }

    const handler = ({ username, color }: { username: string; color: string }) => {
      cbRef.current(username, color);
    };
    colorSocket.on('user-color-changed', handler);
    return () => { colorSocket?.off('user-color-changed', handler); };
  }, [token]);
}

export function emitColorChange(color: string) {
  colorSocket?.emit('update-color', { color });
}
