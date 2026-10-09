import React from "react";

export interface GalleryImageItem {
  id: string;
  url: string;
  title?: string;
  aisle?: string;
  buttonText?: string;
  caption?: string;
  link?: string;
}

interface StoreGalleryProps {
  images?: GalleryImageItem[];
  tagText?: string;
}

export const DEFAULT_GALLERY_IMAGES: GalleryImageItem[] = [
  {
    id: "gal-1",
    url: "https://images.unsplash.com/photo-1535958636474-b021ee887b13?auto=format&fit=crop&q=80&w=800"
  },
  {
    id: "gal-2",
    url: "https://images.unsplash.com/photo-1551024709-8f23befc6f87?auto=format&fit=crop&q=80&w=800"
  },
  {
    id: "gal-3",
    url: "https://images.unsplash.com/photo-1608270111166-5d6e2e50cf16?auto=format&fit=crop&q=80&w=800"
  },
  {
    id: "gal-4",
    url: "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?auto=format&fit=crop&q=80&w=800"
  },
  {
    id: "gal-5",
    url: "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?auto=format&fit=crop&q=80&w=800"
  },
  {
    id: "gal-6",
    url: "https://images.unsplash.com/photo-1516450360452-9312f5e86fc7?auto=format&fit=crop&q=80&w=800"
  }
];

export function StoreGallery({
  images = DEFAULT_GALLERY_IMAGES,
  tagText = "#LASPROMOSDELTIOFELLAS",
}: StoreGalleryProps) {
  // Asegurar exactamente 6 imágenes con formato 4:5
  const displayImages: GalleryImageItem[] = Array.from({ length: 6 }).map((_, index) => {
    return images[index] && images[index].url
      ? images[index]
      : DEFAULT_GALLERY_IMAGES[index] || DEFAULT_GALLERY_IMAGES[0];
  });

  return (
    <section className="w-full px-3 sm:px-6 md:px-8 mb-6 sm:mb-8">
      {/* 
        Línea divisoria simétrica antes de las imágenes con el texto #LASPROMOSDELTIOFELLAS
        situado justo al medio, con una línea a cada lado del texto.
      */}
      <div className="flex items-center gap-3 sm:gap-4 mb-3 sm:mb-4">
        <div className="flex-1 h-px bg-gradient-to-r from-transparent via-neutral-300 to-neutral-200" />
        <span className="text-[10px] sm:text-xs md:text-sm font-black tracking-widest text-[#141414] uppercase shrink-0 px-1 select-none whitespace-nowrap">
          {tagText}
        </span>
        <div className="flex-1 h-px bg-gradient-to-l from-transparent via-neutral-300 to-neutral-200" />
      </div>

      {/* 
        Sin recuadro exterior contenedor, sin textos ni botones, sin puntas redondeadas.
        - 6 imágenes en PC (lg:grid-cols-6 / sm:grid-cols-3)
        - 2 imágenes en celular (grid-cols-2)
        - Relación de aspecto 4:5
        - Solo imagen pura estática (sin lightbox ni clics)
      */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 sm:gap-3">
        {displayImages.map((img, idx) => {
          // Las 2 primeras se ven en celular; las 4 siguientes solo en PC
          const isHiddenOnMobile = idx >= 2;

          return (
            <div
              key={img.id || idx}
              className={`group relative aspect-[4/5] overflow-hidden bg-neutral-100 cursor-default select-none border border-neutral-200 transition-colors ${
                isHiddenOnMobile ? "hidden sm:block" : "block"
              }`}
            >
              {/* Imagen con leve animación de zoom en hover opcional pero estática sin clic */}
              <img
                src={img.url}
                alt=""
                loading="lazy"
                className="w-full h-full object-cover object-center transform scale-100 group-hover:scale-105 transition-transform duration-500 ease-out"
              />
            </div>
          );
        })}
      </div>
    </section>
  );
}
