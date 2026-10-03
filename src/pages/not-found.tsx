import React from "react";

export default function NotFound() {
  return (
    <div className="flex flex-col items-center justify-center min-h-screen bg-background text-foreground p-6">
      <h1 className="text-4xl font-bold mb-2">404</h1>
      <p className="text-lg text-muted-foreground mb-6">Página no encontrada</p>
      <a href="/" className="px-4 py-2 bg-amber-500 text-white rounded-lg font-medium hover:bg-amber-600 transition">
        Volver al inicio
      </a>
    </div>
  );
}
