// src/components/Toast.jsx
import React, { createContext, useContext, useState } from 'react';
import { XCircle } from 'lucide-react';

const ToastContext = createContext();

export const ToastProvider = ({ children }) => {
  const [toasts, setToasts] = useState([]);

  const addToast = (type, message, duration = 4000) => {
    const id = Date.now() + Math.random();
    setToasts((prev) => [...prev, { id, type, message }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, duration);
  };

  return (
    <ToastContext.Provider value={{ addToast }}>
      {children}
      <div className="fixed bottom-4 right-4 z-50 space-y-2">
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`flex items-center gap-2 max-w-xs px-4 py-2 rounded shadow text-sm transition-opacity 
              ${t.type === 'error' ? 'bg-red-100 text-red-800' : 'bg-emerald-100 text-emerald-800'}`}
          >
            <span>{t.message}</span>
            <button onClick={() => setToasts((prev) => prev.filter((toast) => toast.id !== t.id))}>
              <XCircle size={16} className="opacity-70" />
            </button>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
};

export const useToast = () => useContext(ToastContext);
