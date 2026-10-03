import React, { useState } from "react";
import { X } from "lucide-react";

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
  const [selectedImage, setSelectedImage] = useState<GalleryImageItem | null>(null);

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
        <div className="flex-1 h-px bg-gradient-to-r from-transparent via-white/20 to-white/20" />
        <span className="text-[10px] sm:text-xs md:text-sm font-black tracking-widest text-[#ffd025] uppercase shrink-0 px-1 select-none whitespace-nowrap drop-shadow">
          {tagText}
        </span>
        <div className="flex-1 h-px bg-gradient-to-l from-transparent via-white/20 to-white/20" />
      </div>

      {/* 
        Sin recuadro exterior contenedor, sin textos ni botones, sin puntas redondeadas.
        - 6 imágenes en PC (lg:grid-cols-6 / sm:grid-cols-3)
        - 2 imágenes en celular (grid-cols-2)
        - Relación de aspecto 4:5
        - Animación de leve zoom suave en hover
        - Solo imagen pura
      */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 sm:gap-3">
        {displayImages.map((img, idx) => {
          // Las 2 primeras se ven en celular; las 4 siguientes solo en PC
          const isHiddenOnMobile = idx >= 2;

          return (
            <div
              key={img.id || idx}
              onClick={() => setSelectedImage(img)}
              className={`group relative aspect-[4/5] overflow-hidden bg-[#0d0d15] cursor-pointer select-none border border-white/5 hover:border-[#ffd025]/40 transition-colors ${
                isHiddenOnMobile ? "hidden sm:block" : "block"
              }`}
            >
              {/* Imagen con leve animación de zoom en hover */}
              <img
                src={img.url}
                alt=""
                loading="lazy"
                className="w-full h-full object-cover object-center transform scale-100 group-hover:scale-108 transition-transform duration-500 ease-out"
              />
            </div>
          );
        })}
      </div>

      {/* Modal Lightbox minimalista: solo la imagen ampliada en 4:5 sin bordes redondeados ni texto */}
      {selectedImage && (
        <div
          className="fixed inset-0 z-50 bg-black/95 flex items-center justify-center p-3 sm:p-6 animate-fade-in"
          onClick={() => setSelectedImage(null)}
        >
          <div
            className="relative max-w-md sm:max-w-lg w-full bg-black shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <button
              onClick={() => setSelectedImage(null)}
              className="absolute top-2 right-2 z-10 w-8 h-8 bg-black/80 text-white flex items-center justify-center hover:bg-[#ffd025] hover:text-black transition-colors"
              title="Cerrar"
            >
              <X size={18} />
            </button>
            <div className="relative aspect-[4/5] w-full bg-black">
              <img
                src={selectedImage.url}
                alt=""
                className="w-full h-full object-cover"
              />
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
