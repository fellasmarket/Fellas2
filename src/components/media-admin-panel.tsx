import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  Upload,
  Image as ImageIcon,
  Trash2,
  Copy,
  Check,
  Search,
  RefreshCw,
  FolderOpen,
  Sparkles,
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Tag,
  Flame,
  Wine,
  Beer,
  GlassWater,
  Cigarette,
  Cookie,
  Candy,
  Package,
  Layers,
  Bot,
  Zap,
  Droplets,
  Star,
  Plus,
  X,
  ShoppingBag,
  CheckCircle,
  CheckCircle2,
} from "lucide-react";
import { uploadImagesBatch } from "../lib/image-batch-uploader";

export interface MediaItem {
  id: string;
  name: string;
  url: string;
  createdAt: string;
  size?: number;
  category?: string;
  aiDetectedTitle?: string;
}

export interface MediaProduct {
  id: number;
  name: string;
  price: number;
  image?: string;
  images?: string[];
  category?: string;
  aisle?: string;
  subcategory?: string;
}

interface MediaAdminPanelProps {
  onMediaDeleted?: () => void;
  products?: any[];
  onProductUpdated?: () => void;
}

export const CATEGORY_ICONS: Record<string, React.ReactNode> = {
  "Bebidas Alcohólicas": <GlassWater size={16} className="text-purple-400" />,
  "Piscos": <Wine size={16} className="text-[#ffd025]" />,
  "Cervezas": <Beer size={16} className="text-yellow-400" />,
  "Vinos & Espumantes": <Wine size={16} className="text-red-400" />,
  "Destilados & Licores": <GlassWater size={16} className="text-cyan-400" />,
  "Energizantes": <Zap size={16} className="text-lime-400" />,
  "Bebidas & Gaseosas": <GlassWater size={16} className="text-emerald-400" />,
  "Jugos": <Droplets size={16} className="text-amber-400" />,
  "Aguas": <Droplets size={16} className="text-sky-400" />,
  "Cigarros & Tabacos": <Cigarette size={16} className="text-orange-400" />,
  "Snacks & Salados": <Cookie size={16} className="text-amber-500" />,
  "Dulces & Chocolates": <Candy size={16} className="text-pink-400" />,
  "Promociones & Packs": <Flame size={16} className="text-amber-400" />,
  "Hielo & Abarrotes": <Package size={16} className="text-blue-300" />,
  "Otros / General": <FolderOpen size={16} className="text-gray-400" />,
};

export const ORDERED_SECTIONS = [
  "Bebidas Alcohólicas",
  "Piscos",
  "Cervezas",
  "Vinos & Espumantes",
  "Destilados & Licores",
  "Energizantes",
  "Bebidas & Gaseosas",
  "Jugos",
  "Aguas",
  "Cigarros & Tabacos",
  "Snacks & Salados",
  "Dulces & Chocolates",
  "Promociones & Packs",
  "Hielo & Abarrotes",
  "Otros / General",
];

const normalizeText = (text: string) =>
  (text || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();

interface ImageAssignedBadgesProps {
  imageUrl: string;
  products: MediaProduct[];
  onOpenAssignModal: () => void;
  onUnassign: (imgUrl: string, productId: number) => Promise<void>;
}

export const ImageAssignedBadges: React.FC<ImageAssignedBadgesProps> = ({
  imageUrl,
  products,
  onOpenAssignModal,
  onUnassign,
}) => {
  const [unassigningId, setUnassigningId] = useState<number | null>(null);

  const assignedProducts = useMemo(() => {
    if (!imageUrl) return [];
    return products.filter((p) => {
      if (p.image === imageUrl) return true;
      if (Array.isArray(p.images) && p.images.includes(imageUrl)) return true;
      return false;
    });
  }, [products, imageUrl]);

  const handleRemove = async (productId: number, e: React.MouseEvent) => {
    e.stopPropagation();
    try {
      setUnassigningId(productId);
      await onUnassign(imageUrl, productId);
    } finally {
      setUnassigningId(null);
    }
  };

  return (
    <div className="w-full space-y-2 pt-1">
      {/* Badges de productos asignados */}
      {assignedProducts.length > 0 ? (
        <div className="space-y-1">
          <div className="flex items-center justify-between text-[9px] text-gray-400 font-bold uppercase tracking-wider">
            <span>Vinculada a:</span>
            <span className="text-[#ffd025] font-mono font-bold bg-[#ffd025]/10 px-1.5 py-0.2 rounded-full">
              {assignedProducts.length}
            </span>
          </div>
          <div className="flex flex-wrap gap-1 max-h-20 overflow-y-auto">
            {assignedProducts.map((p) => {
              const isPrimary = p.image === imageUrl;
              return (
                <span
                  key={p.id}
                  className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-lg bg-[#ffd025]/15 border border-[#ffd025]/30 text-[#ffd025] text-[9.5px] font-bold max-w-full"
                  title={`Asignada a ${p.name} ${isPrimary ? '(Foto Principal)' : '(Galería)'}`}
                >
                  <Tag size={9} className="shrink-0" />
                  <span className="truncate max-w-[85px] sm:max-w-[110px]">{p.name}</span>
                  {isPrimary && (
                    <span className="text-[7.5px] bg-[#ffd025] text-black px-1 rounded font-black shrink-0">
                      1ª
                    </span>
                  )}
                  <button
                    type="button"
                    onClick={(e) => handleRemove(p.id, e)}
                    disabled={unassigningId === p.id}
                    className="hover:text-red-400 text-gray-400 p-0.5 transition cursor-pointer"
                    title="Desvincular imagen de este producto"
                  >
                    <X size={9} />
                  </button>
                </span>
              );
            })}
          </div>
        </div>
      ) : (
        <p className="text-[9px] text-gray-500 flex items-center gap-1 py-0.5">
          <span className="w-1.5 h-1.5 rounded-full bg-gray-600 inline-block shrink-0" />
          <span className="truncate">Sin producto asignado</span>
        </p>
      )}

      {/* Botón cómodo y espacioso para abrir el modal de asignación */}
      <button
        type="button"
        onClick={onOpenAssignModal}
        className="w-full py-2 px-2.5 rounded-xl bg-[#ffd025]/15 hover:bg-[#ffd025] text-[#ffd025] hover:text-black border border-[#ffd025]/40 font-black text-[10.5px] sm:text-xs flex items-center justify-center gap-1.5 transition-all shadow-md active:scale-95 cursor-pointer group"
        title="Abrir buscador predictivo para asignar a producto"
      >
        <Search size={13} className="text-[#ffd025] group-hover:text-black shrink-0" strokeWidth={2.5} />
        <span className="truncate">Asignar a Producto</span>
      </button>
    </div>
  );
};

interface AssignProductModalProps {
  image: MediaItem;
  products: MediaProduct[];
  onClose: () => void;
  onAssign: (imgUrl: string, productId: number, asPrimary: boolean) => Promise<void>;
  onUnassign: (imgUrl: string, productId: number) => Promise<void>;
}

export const AssignProductModal: React.FC<AssignProductModalProps> = ({
  image,
  products,
  onClose,
  onAssign,
  onUnassign,
}) => {
  const [query, setQuery] = useState("");
  const [assigningId, setAssigningId] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const timer = setTimeout(() => {
      inputRef.current?.focus();
    }, 150);
    return () => clearTimeout(timer);
  }, []);

  const assignedProducts = useMemo(() => {
    if (!image.url) return [];
    return products.filter((p) => {
      if (p.image === image.url) return true;
      if (Array.isArray(p.images) && p.images.includes(image.url)) return true;
      return false;
    });
  }, [products, image.url]);

  const matchingProducts = useMemo(() => {
    const q = normalizeText(query);
    if (!q) {
      return products.slice(0, 30);
    }
    return products
      .filter((p) => {
        const n = normalizeText(p.name);
        const c = normalizeText(p.category || "");
        const a = normalizeText(p.aisle || "");
        const s = normalizeText(p.subcategory || "");
        return n.includes(q) || c.includes(q) || a.includes(q) || s.includes(q);
      })
      .slice(0, 50);
  }, [products, query]);

  const handleSelect = async (p: MediaProduct, asPrimary: boolean) => {
    try {
      setAssigningId(p.id);
      await onAssign(image.url, p.id, asPrimary);
    } finally {
      setAssigningId(null);
    }
  };

  const handleRemove = async (productId: number) => {
    try {
      setAssigningId(productId);
      await onUnassign(image.url, productId);
    } finally {
      setAssigningId(null);
    }
  };

  return (
    <div
      onClick={onClose}
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-black/85 backdrop-blur-md animate-fade-in"
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="relative max-w-2xl w-full bg-[#131322] border-2 border-[#ffd025]/60 rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh] animate-scale-in text-left"
      >
        {/* Encabezado con la foto seleccionada y título amplio */}
        <div className="p-4 sm:p-5 bg-[#171728] border-b border-white/10 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3.5 min-w-0">
            <div className="w-14 h-14 sm:w-16 sm:h-16 rounded-2xl bg-black/70 border border-white/10 overflow-hidden flex items-center justify-center shrink-0 p-1">
              <img
                src={image.url}
                alt={image.name}
                className="w-full h-full object-contain"
                onError={(e) => {
                  (e.currentTarget as HTMLImageElement).src =
                    "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                }}
              />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="px-2 py-0.5 rounded-md bg-[#ffd025]/20 text-[#ffd025] text-[10px] font-black uppercase tracking-wider">
                  {image.category || "General"}
                </span>
                <span className="text-xs text-gray-400 font-mono hidden sm:inline">
                  {assignedProducts.length > 0 ? `${assignedProducts.length} producto(s) vinculado(s)` : "Sin vincular"}
                </span>
              </div>
              <h3 className="text-base sm:text-lg font-black text-white uppercase tracking-wider truncate">
                Asignar Imagen a Producto
              </h3>
              <p className="text-[11px] text-gray-400 truncate max-w-md" title={image.name}>
                Archivo: {image.name}
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-2.5 rounded-2xl bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white transition cursor-pointer shrink-0 border border-white/10"
            title="Cerrar ventana"
          >
            <X size={18} />
          </button>
        </div>

        {/* Resumen de productos que ya tienen esta imagen asignada */}
        {assignedProducts.length > 0 && (
          <div className="px-4 sm:px-5 py-3 bg-[#ffd025]/10 border-b border-[#ffd025]/20 flex flex-wrap items-center gap-2">
            <span className="text-[11px] font-bold text-[#ffd025] uppercase tracking-wider flex items-center gap-1 shrink-0">
              <CheckCircle2 size={13} /> Asignada actualmente a:
            </span>
            <div className="flex flex-wrap gap-1.5 flex-1">
              {assignedProducts.map((p) => {
                const isPrimary = p.image === image.url;
                const isBusy = assigningId === p.id;
                return (
                  <span
                    key={p.id}
                    className="inline-flex items-center gap-1.5 px-2 py-1 rounded-xl bg-black/60 border border-[#ffd025]/40 text-white text-xs font-bold"
                  >
                    <span>{p.name}</span>
                    <span className="text-[9px] px-1.5 py-0.2 rounded font-black uppercase bg-[#ffd025] text-black">
                      {isPrimary ? "1ª Principal" : "Galería"}
                    </span>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleRemove(p.id)}
                      className="text-gray-400 hover:text-red-400 p-0.5 transition cursor-pointer"
                      title="Desvincular esta foto del producto"
                    >
                      <X size={12} />
                    </button>
                  </span>
                );
              })}
            </div>
          </div>
        )}

        {/* Barra de Búsqueda Predictiva GRANDE Y ESPACIOSA */}
        <div className="p-4 sm:p-5 border-b border-white/10 bg-[#0e0e1a] space-y-2">
          <div className="relative w-full">
            <Search size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-[#ffd025]" />
            <input
              ref={inputRef}
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Escribe el nombre del producto (ej: Mistral, Heineken, Lays, Coca-Cola)..."
              className="w-full pl-11 pr-11 py-3.5 bg-[#171728] border-2 border-[#ffd025]/40 focus:border-[#ffd025] rounded-2xl text-white placeholder-gray-400 text-xs sm:text-sm font-semibold shadow-inner focus:outline-none transition-all"
            />
            {query && (
              <button
                type="button"
                onClick={() => setQuery("")}
                className="absolute right-3.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-white p-1 rounded-xl hover:bg-white/10 transition"
              >
                <X size={16} />
              </button>
            )}
          </div>

          <div className="flex items-center justify-between text-[11px] text-gray-400 px-1">
            <span>
              {query
                ? `Coincidencias encontradas: ${matchingProducts.length}`
                : `Sugerencias de productos de la tienda (${matchingProducts.length})`}
            </span>
            <span className="text-[#ffd025] font-bold hidden sm:inline">
              Haz clic en &quot;⭐ Principal&quot; o &quot;➕ Galería&quot;
            </span>
          </div>
        </div>

        {/* Lista de coincidencias amplia, cómoda y sin solaparse */}
        <div className="p-3 sm:p-5 overflow-y-auto space-y-2.5 flex-1 max-h-[50vh] bg-[#0c0c14]">
          {matchingProducts.length === 0 ? (
            <div className="py-12 px-4 text-center text-gray-400 space-y-2">
              <ShoppingBag size={36} className="mx-auto text-gray-600 mb-2" />
              <p className="text-sm font-bold text-white">
                No se encontró ningún producto con &quot;{query}&quot;
              </p>
              <p className="text-xs text-gray-500">
                Verifica el nombre o prueba buscando por categoría (ej: Cerveza, Pisco, Snack, Bebida).
              </p>
            </div>
          ) : (
            matchingProducts.map((p) => {
              const isAssigned = p.image === image.url || (Array.isArray(p.images) && p.images.includes(image.url));
              const isPrimary = p.image === image.url;
              const isBusy = assigningId === p.id;

              return (
                <div
                  key={p.id}
                  className={`p-3 sm:p-3.5 rounded-2xl border transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                    isAssigned
                      ? "bg-[#18182c] border-[#ffd025]/40 shadow-lg"
                      : "bg-[#141422] border-white/5 hover:border-white/20 hover:bg-[#181828]"
                  }`}
                >
                  {/* Foto, Nombre y Datos del Producto */}
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="w-12 h-12 rounded-xl bg-black/60 border border-white/10 overflow-hidden flex items-center justify-center shrink-0 p-1">
                      {p.image ? (
                        <img
                          src={p.image}
                          alt={p.name}
                          className="w-full h-full object-contain"
                          onError={(e) => {
                            (e.currentTarget as HTMLImageElement).src =
                              "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                          }}
                        />
                      ) : (
                        <ShoppingBag size={18} className="text-gray-500" />
                      )}
                    </div>
                    <div className="min-w-0">
                      <h4 className="text-xs sm:text-sm font-black text-white truncate" title={p.name}>
                        {p.name}
                      </h4>
                      <div className="flex flex-wrap items-center gap-2 text-[10.5px] text-gray-400 mt-0.5">
                        <span className="text-[#ffd025] font-mono font-bold text-xs">
                          ${(p.price || 0).toLocaleString("es-CL")}
                        </span>
                        <span>•</span>
                        <span className="px-2 py-0.5 rounded-md bg-white/5 text-gray-300 font-semibold text-[10px]">
                          {p.aisle || p.category || "General"}
                        </span>
                        {p.subcategory && (
                          <span className="text-gray-500 text-[10px] hidden sm:inline">
                            / {p.subcategory}
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Botones de acción cómodos, grandes y claros */}
                  <div className="flex items-center gap-2 shrink-0 self-end sm:self-center">
                    {isAssigned ? (
                      <div className="flex items-center gap-2">
                        <span className="px-2.5 py-1.5 rounded-xl bg-emerald-500/20 border border-emerald-500/30 text-emerald-300 text-xs font-bold flex items-center gap-1.5">
                          <Check size={14} /> Asignada {isPrimary ? "(1ª)" : "(Galería)"}
                        </span>
                        {!isPrimary && (
                          <button
                            type="button"
                            disabled={isBusy}
                            onClick={() => handleSelect(p, true)}
                            className="px-3 py-1.5 rounded-xl bg-[#ffd025]/20 hover:bg-[#ffd025] text-[#ffd025] hover:text-black text-xs font-bold transition flex items-center gap-1.5 cursor-pointer border border-[#ffd025]/40"
                            title="Cambiar a foto principal de portada"
                          >
                            <Star size={12} />
                            <span>Hacer 1ª Principal</span>
                          </button>
                        )}
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleRemove(p.id)}
                          className="px-3 py-1.5 rounded-xl bg-red-500/20 hover:bg-red-500 text-red-300 hover:text-white text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-red-500/30"
                          title="Desvincular imagen de este producto"
                        >
                          <X size={12} />
                          <span>Desvincular</span>
                        </button>
                      </div>
                    ) : (
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleSelect(p, true)}
                          className="px-3.5 py-2 rounded-xl bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] hover:scale-105 active:scale-95 text-black font-black uppercase text-[11px] tracking-wider transition-all flex items-center gap-1.5 shadow-md shadow-[#ffd025]/20 cursor-pointer"
                          title="Establecer como la foto principal que se ve en la tienda"
                        >
                          <Star size={13} strokeWidth={2.5} />
                          <span>⭐ Foto Principal</span>
                        </button>

                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleSelect(p, false)}
                          className="px-3 py-2 rounded-xl bg-white/10 hover:bg-white/20 active:scale-95 text-white font-bold text-[11px] transition-all flex items-center gap-1.5 border border-white/10 cursor-pointer"
                          title="Añadir como foto adicional en el carrusel del producto"
                        >
                          <Plus size={13} strokeWidth={2.5} />
                          <span>➕ Galería</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer del Modal */}
        <div className="p-4 sm:p-5 bg-[#171728] border-t border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <p className="text-[11px] text-gray-400">
            Tip: Puedes asignar la misma imagen a más de un producto. Todos los cambios se guardan al instante.
          </p>
          <button
            type="button"
            onClick={onClose}
            className="w-full sm:w-auto px-6 py-2.5 rounded-xl bg-white/10 hover:bg-white/20 text-white font-black uppercase text-xs tracking-wider transition cursor-pointer"
          >
            Listo / Cerrar
          </button>
        </div>
      </div>
    </div>
  );
};

export const MediaAdminPanel: React.FC<MediaAdminPanelProps> = ({
  onMediaDeleted,
  products: propsProducts,
  onProductUpdated,
}) => {
  const [images, setImages] = useState<MediaItem[]>([]);
  const [products, setProducts] = useState<MediaProduct[]>(propsProducts || []);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState("");
  const [uploadPercent, setUploadPercent] = useState(0);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedSectionFilter, setSelectedSectionFilter] = useState<string>("all");
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [previewImage, setPreviewImage] = useState<MediaItem | null>(null);
  const [classifyingAI, setClassifyingAI] = useState(false);
  const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
  const [assigningImage, setAssigningImage] = useState<MediaItem | null>(null);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((cur) => (cur === msg ? null : cur));
    }, 3200);
  };

  const fetchProducts = async () => {
    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/products`);
      if (res.ok) {
        const data = await res.json();
        setProducts(data || []);
      }
    } catch (err) {
      console.error("Error fetching products in MediaAdminPanel:", err);
    }
  };

  useEffect(() => {
    if (propsProducts && propsProducts.length > 0) {
      setProducts(propsProducts);
    } else {
      fetchProducts();
    }
  }, [propsProducts]);

  const handleAssignImageToProduct = async (
    imgUrl: string,
    productId: number,
    asPrimary: boolean
  ) => {
    const prod = products.find((p) => p.id === productId);
    if (!prod) return;

    const currentImages = Array.isArray(prod.images) && prod.images.length > 0
      ? prod.images
      : (prod.image ? [prod.image] : []);

    let newImages: string[];
    let newPrimary: string;

    if (asPrimary) {
      newPrimary = imgUrl;
      newImages = [imgUrl, ...currentImages.filter((u) => u !== imgUrl)];
    } else {
      newPrimary = prod.image || imgUrl;
      newImages = currentImages.includes(imgUrl)
        ? currentImages
        : [...currentImages, imgUrl];
    }

    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/products/${productId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          image: newPrimary,
          images: newImages,
        }),
      });

      if (res.ok) {
        setProducts((prev) =>
          prev.map((p) =>
            p.id === productId ? { ...p, image: newPrimary, images: newImages } : p
          )
        );
        onProductUpdated?.();
        showToast(
          asPrimary
            ? `✓ Foto principal asignada a "${prod.name}"`
            : `✓ Foto añadida a la galería de "${prod.name}"`
        );
      } else {
        alert("No se pudo asignar la imagen al producto.");
      }
    } catch (err) {
      console.error("Error assigning image to product:", err);
      alert("Error de conexión al asignar imagen.");
    }
  };

  const handleUnassignImageFromProduct = async (imgUrl: string, productId: number) => {
    const prod = products.find((p) => p.id === productId);
    if (!prod) return;

    const currentImages = Array.isArray(prod.images) && prod.images.length > 0
      ? prod.images
      : (prod.image ? [prod.image] : []);

    const newImages = currentImages.filter((u) => u !== imgUrl);
    const newPrimary = prod.image === imgUrl ? (newImages[0] || "") : prod.image;

    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/products/${productId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          image: newPrimary,
          images: newImages,
        }),
      });

      if (res.ok) {
        setProducts((prev) =>
          prev.map((p) =>
            p.id === productId ? { ...p, image: newPrimary, images: newImages } : p
          )
        );
        onProductUpdated?.();
        showToast(`Foto desvinculada de "${prod.name}"`);
      }
    } catch (err) {
      console.error("Error unassigning image:", err);
    }
  };

  const fetchImages = async () => {
    try {
      setLoading(true);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library`);
      if (res.ok) {
        const data = await res.json();
        setImages(data.images || []);
      }
    } catch (err) {
      console.error("Error fetching media library:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchImages();
  }, []);

  const handleBatchUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;

    const filesArray = Array.from(fileList);
    setUploading(true);
    setUploadPercent(0);
    setUploadProgress(`Iniciando carga de ${filesArray.length} imágenes...`);

    try {
      const { successCount, failedCount } = await uploadImagesBatch(
        filesArray,
        (p) => {
          setUploadPercent(p.percent);
          setUploadProgress(
            `Procesando y subiendo: ${p.current} de ${p.total} (${p.percent}%) — ${p.fileName}`
          );
        }
      );

      await fetchImages();

      if (failedCount > 0) {
        alert(`Se subieron ${successCount} imágenes con éxito (${failedCount} no se pudieron procesar).`);
      }
    } catch (err) {
      console.error("Error in batch upload:", err);
      alert("Hubo un error al subir el grupo de imágenes.");
    } finally {
      setUploading(false);
      setUploadProgress("");
      setUploadPercent(0);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const handleClassifyAI = async () => {
    try {
      setClassifyingAI(true);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/classify-ai`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ forceAll: true }),
      });
      if (res.ok) {
        const data = await res.json();
        setImages(data.images || []);
        alert(`¡Organización con IA completada! Se analizaron y organizaron ${data.updatedCount || images.length} imágenes en sus respectivas secciones.`);
      } else {
        alert("Error al organizar con IA.");
      }
    } catch (err) {
      console.error("Error organizing with AI:", err);
      alert("Error de conexión al ejecutar la organización por IA.");
    } finally {
      setClassifyingAI(false);
    }
  };

  const handleUpdateCategory = async (id: string, newCategory: string) => {
    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: newCategory }),
      });
      if (res.ok) {
        setImages((prev) =>
          prev.map((img) => (img.id === id ? { ...img, category: newCategory } : img))
        );
      }
    } catch (err) {
      console.error("Error updating image category:", err);
    }
  };

  const handleDeleteImage = async (id: string) => {
    if (!confirm("¿Deseas eliminar permanentemente esta imagen del archivo?")) return;

    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/${id}`, {
        method: "DELETE",
      });
      if (res.ok) {
        setImages((prev) => prev.filter((img) => img.id !== id));
        if (previewImage?.id === id) setPreviewImage(null);
        onMediaDeleted?.();
      }
    } catch (err) {
      console.error("Error deleting image:", err);
    }
  };

  const handleDeleteAllImages = async () => {
    if (images.length === 0) return;
    const confirmed = confirm(
      `¿Estás seguro de que deseas eliminar permanentemente TODAS las ${images.length} imágenes del archivo?\n\nEsta acción borrará todas las fotos subidas previamente del almacenamiento y las removerá de los productos que las tengan asignadas.`
    );
    if (!confirmed) return;

    try {
      setLoading(true);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/all`, {
        method: "DELETE",
      });
      if (res.ok) {
        setImages([]);
        setPreviewImage(null);
        onMediaDeleted?.();
        alert("Se han eliminado todas las imágenes del archivo y se han desvinculado de los productos con éxito.");
      } else {
        alert("Error al eliminar las imágenes del archivo.");
      }
    } catch (err) {
      console.error("Error deleting all images:", err);
      alert("Error de conexión al eliminar las imágenes.");
    } finally {
      setLoading(false);
    }
  };

  const handleCopyUrl = (url: string, id: string) => {
    navigator.clipboard.writeText(url);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const toggleSection = (secName: string) => {
    setCollapsedSections((prev) => ({
      ...prev,
      [secName]: !prev[secName],
    }));
  };

  const toggleAllSections = (collapse: boolean) => {
    const newState: Record<string, boolean> = {};
    ORDERED_SECTIONS.forEach((s) => {
      newState[s] = collapse;
    });
    setCollapsedSections(newState);
  };

  // Group images by section
  const { groupedImages, sectionCounts, allFilteredCount } = useMemo(() => {
    const counts: Record<string, number> = {};
    ORDERED_SECTIONS.forEach((s) => (counts[s] = 0));

    const q = normalizeText(searchQuery);

    const filtered = images.filter((img) => {
      if (!q) return true;
      const matchName = normalizeText(img.name).includes(q);
      const matchCat = normalizeText(img.category || "").includes(q);
      const matchTitle = normalizeText(img.aiDetectedTitle || "").includes(q);
      const matchAssigned = products.some(
        (p) =>
          (p.image === img.url || (Array.isArray(p.images) && p.images.includes(img.url))) &&
          (normalizeText(p.name).includes(q) ||
            normalizeText(p.aisle || "").includes(q) ||
            normalizeText(p.category || "").includes(q))
      );
      return matchName || matchCat || matchTitle || matchAssigned;
    });

    const grouped: Record<string, MediaItem[]> = {};
    ORDERED_SECTIONS.forEach((s) => (grouped[s] = []));

    filtered.forEach((img) => {
      const cat = img.category && ORDERED_SECTIONS.includes(img.category) ? img.category : "Otros / General";
      counts[cat] = (counts[cat] || 0) + 1;
      grouped[cat].push(img);
    });

    return {
      groupedImages: grouped,
      sectionCounts: counts,
      allFilteredCount: filtered.length,
    };
  }, [images, searchQuery, products]);

  return (
    <div className="space-y-6 max-w-6xl mx-auto animate-fade-in text-left">
      {/* Top Banner Card */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-7 rounded-3xl border border-white/10 flex flex-col md:flex-row md:items-center justify-between gap-5 shadow-xl shadow-black/60">
        <div>
          <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
            <Sparkles size={14} /> Archivo Central Inteligente con IA
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-white uppercase tracking-wider">
            Galería y Archivo de Medios
          </h2>
          <p className="text-xs text-gray-400 mt-1 max-w-xl">
            Sube múltiples imágenes en lote. La Inteligencia Artificial las clasifica automáticamente por tipo de producto (Piscos, Cervezas, Vinos, Promos, Cigarros, Snacks, etc.) en menús plegables para que las encuentres al instante.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={fetchImages}
            disabled={loading || classifyingAI}
            className="p-3 rounded-2xl bg-white/5 hover:bg-white/10 text-gray-300 transition-colors border border-white/10 cursor-pointer"
            title="Recargar archivo"
          >
            <RefreshCw size={18} className={loading ? "animate-spin text-[#ffd025]" : ""} />
          </button>

          {images.length > 0 && (
            <button
              type="button"
              onClick={handleClassifyAI}
              disabled={loading || classifyingAI || uploading}
              className="px-4 py-3.5 rounded-2xl bg-gradient-to-r from-purple-600/30 to-indigo-600/30 hover:from-purple-600/50 hover:to-indigo-600/50 text-purple-300 hover:text-white border border-purple-500/40 transition-all font-bold text-xs flex items-center gap-2 cursor-pointer shrink-0 shadow-lg shadow-purple-950/40"
              title="Analizar y organizar todas las imágenes en secciones mediante IA"
            >
              <Bot size={17} className={classifyingAI ? "animate-bounce text-[#ffd025]" : "text-purple-400"} />
              <span>{classifyingAI ? "Analizando con IA..." : "Organizar con IA"}</span>
            </button>
          )}

          {images.length > 0 && (
            <button
              type="button"
              onClick={handleDeleteAllImages}
              disabled={loading || uploading || classifyingAI}
              className="px-4 py-3.5 rounded-2xl bg-red-600/20 hover:bg-red-600/30 text-red-400 hover:text-red-300 border border-red-500/30 transition-all font-bold text-xs flex items-center gap-2 cursor-pointer shrink-0 shadow-lg shadow-red-950/40"
              title="Eliminar todas las imágenes subidas del archivo de medios"
            >
              <Trash2 size={16} />
              <span>Borrar Todas ({images.length})</span>
            </button>
          )}

          <label
            className={`cursor-pointer px-6 py-3.5 rounded-2xl font-black uppercase text-xs tracking-wider transition-all flex items-center justify-center gap-2 shadow-xl shadow-[#ffd025]/20 shrink-0 ${
              uploading
                ? "bg-gray-700 text-gray-400 cursor-not-allowed"
                : "bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] hover:scale-[1.02] active:scale-[0.98]"
            }`}
          >
            <Upload size={18} strokeWidth={2.5} />
            <span>{uploading ? "Subiendo lote..." : "Subir Grupo de Imágenes"}</span>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept="image/*"
              disabled={uploading}
              onChange={handleBatchUpload}
              className="hidden"
            />
          </label>
        </div>
      </div>

      {/* Upload Progress Alert & Bar */}
      {uploading && (
        <div className="p-4 sm:p-5 rounded-2xl bg-[#ffd025]/15 border border-[#ffd025]/30 space-y-2.5 animate-fade-in shadow-xl shadow-[#ffd025]/10">
          <div className="flex items-center justify-between gap-3 text-xs sm:text-sm font-bold text-[#ffd025]">
            <span className="flex items-center gap-2 truncate">
              <RefreshCw size={16} className="animate-spin text-[#ffd025] shrink-0" />
              <span className="truncate">{uploadProgress}</span>
            </span>
            <span className="font-mono text-sm font-black shrink-0">{uploadPercent}%</span>
          </div>
          <div className="w-full h-2.5 rounded-full bg-black/50 overflow-hidden border border-white/10">
            <div
              className="h-full bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] transition-all duration-300 rounded-full"
              style={{ width: `${uploadPercent}%` }}
            />
          </div>
        </div>
      )}

      {/* ACCESO RÁPIDO / SELECTOR DE SECCIONES SUPERIOR */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-3xl border border-white/10 space-y-3 shadow-xl">
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
          {/* Search Box */}
          <div className="relative w-full sm:w-80">
            <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
            <input
              type="text"
              placeholder="Buscar por nombre o tipo..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2.5 bg-[#0e0e16] border border-white/10 rounded-xl text-white text-xs focus:border-[#ffd025] focus:outline-none"
            />
          </div>

          {/* Master Accordion Controls & Stats */}
          <div className="flex items-center gap-2 justify-between sm:justify-end">
            <div className="text-xs text-gray-400 font-mono flex items-center gap-2 mr-2">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse" />
              <span>{allFilteredCount} fotos</span>
            </div>

            <button
              type="button"
              onClick={() => toggleAllSections(false)}
              className="px-2.5 py-1.5 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 hover:text-white text-[11px] font-bold transition border border-white/10 cursor-pointer"
            >
              Expandir todo
            </button>
            <button
              type="button"
              onClick={() => toggleAllSections(true)}
              className="px-2.5 py-1.5 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 hover:text-white text-[11px] font-bold transition border border-white/10 cursor-pointer"
            >
              Plegar todo
            </button>
          </div>
        </div>

        {/* Barra de Acceso Rápido Selector de Secciones (Scroll Horizontal) */}
        <div className="pt-2 border-t border-white/10 flex items-center gap-1.5 sm:gap-2 overflow-x-auto pb-1 scrollbar-none">
          <button
            type="button"
            onClick={() => setSelectedSectionFilter("all")}
            className={`px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider whitespace-nowrap transition flex items-center gap-1.5 cursor-pointer border ${
              selectedSectionFilter === "all"
                ? "bg-[#ffd025] text-black border-[#ffd025] shadow-md shadow-[#ffd025]/20 font-black"
                : "bg-white/5 text-gray-400 hover:text-white border-white/10"
            }`}
          >
            <Layers size={13} />
            <span>Todas ({images.length})</span>
          </button>

          {ORDERED_SECTIONS.map((secName) => {
            const count = sectionCounts[secName] || 0;
            if (count === 0 && selectedSectionFilter !== secName) return null;
            const isSelected = selectedSectionFilter === secName;
            return (
              <button
                key={secName}
                type="button"
                onClick={() => {
                  setSelectedSectionFilter(secName);
                  // Ensure section is expanded
                  setCollapsedSections((prev) => ({ ...prev, [secName]: false }));
                }}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition flex items-center gap-1.5 cursor-pointer border ${
                  isSelected
                    ? "bg-[#ffd025] text-black border-[#ffd025] shadow-md shadow-[#ffd025]/20 font-black"
                    : "bg-white/5 text-gray-300 hover:text-white border-white/10 hover:border-white/20"
                }`}
              >
                {CATEGORY_ICONS[secName] || <Tag size={13} />}
                <span>{secName}</span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-mono font-bold ${isSelected ? "bg-black/30 text-black" : "bg-white/10 text-gray-300"}`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {/* SECCIONES PLEGABLES (MENÚS PLEGABLES / ACCORDION) */}
      {loading && images.length === 0 ? (
        <div className="p-16 rounded-3xl bg-[#13131f]/60 border border-white/10 flex flex-col items-center justify-center text-gray-400">
          <RefreshCw size={36} className="animate-spin text-[#ffd025] mb-3" />
          <p className="text-sm font-bold">Cargando y clasificando archivo de medios...</p>
        </div>
      ) : allFilteredCount === 0 ? (
        <div className="p-16 rounded-3xl bg-[#13131f]/60 border border-white/10 flex flex-col items-center justify-center text-center">
          <div className="w-20 h-20 rounded-3xl bg-white/5 border border-white/10 flex items-center justify-center mb-4 text-[#ffd025]">
            <FolderOpen size={36} />
          </div>
          <h3 className="text-lg font-black text-white uppercase tracking-wider mb-2">
            {searchQuery ? "No hay imágenes que coincidan con la búsqueda" : "No has subido imágenes todavía"}
          </h3>
          <p className="text-xs text-gray-400 max-w-md mb-6">
            {searchQuery
              ? "Prueba buscando con otra palabra o seleccionando 'Todas' en el acceso rápido."
              : "Selecciona varias fotos de tus productos para que la IA las organice automáticamente por secciones."}
          </p>
          {!searchQuery && (
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-7 py-3 bg-[#ffd025] hover:bg-[#ffe066] text-black font-black uppercase text-xs tracking-wider rounded-xl transition shadow-lg shadow-[#ffd025]/20 cursor-pointer"
            >
              Seleccionar Imágenes Ahora
            </button>
          )}
        </div>
      ) : (
        <div className="space-y-4">
          {ORDERED_SECTIONS.map((secName) => {
            const secImages = groupedImages[secName] || [];
            if (secImages.length === 0) return null;
            if (selectedSectionFilter !== "all" && selectedSectionFilter !== secName) return null;

            const isCollapsed = Boolean(collapsedSections[secName]);

            return (
              <div
                key={secName}
                id={`section-${secName.replace(/\s+/g, "-")}`}
                className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl transition-all"
              >
                {/* ENCABEZADO PLEGABLE DE LA SECCIÓN */}
                <button
                  type="button"
                  onClick={() => toggleSection(secName)}
                  className="w-full px-5 sm:px-6 py-4 flex items-center justify-between bg-[#171728]/80 hover:bg-[#1a1a2e] transition cursor-pointer border-b border-white/10 text-left"
                >
                  <div className="flex items-center gap-3">
                    <div className="w-9 h-9 rounded-xl bg-black/40 border border-white/10 flex items-center justify-center shadow-inner">
                      {CATEGORY_ICONS[secName] || <FolderOpen size={16} />}
                    </div>
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="text-sm sm:text-base font-black text-white uppercase tracking-wider">
                          {secName}
                        </h3>
                        <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#ffd025]/20 text-[#ffd025] font-mono font-bold">
                          {secImages.length} {secImages.length === 1 ? "foto" : "fotos"}
                        </span>
                      </div>
                      <p className="text-[10.5px] text-gray-400">
                        {secName === "Promociones & Packs"
                          ? "Combos armados (pisco+bebida+hielo, whisky+bebida, ofertas combinadas)"
                          : secName === "Piscos"
                          ? "Piscos nacionales e internacionales"
                          : secName === "Cervezas"
                          ? "Latas, botellas y six-packs"
                          : secName === "Vinos & Espumantes"
                          ? "Tintos, blancos, espumantes y champagne"
                          : secName === "Destilados & Licores"
                          ? "Whisky, Ron, Vodka, Gin, Tequila y Licores"
                          : `Sección de ${secName}`}
                      </p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 text-gray-400">
                    <span className="text-[11px] font-bold hidden sm:inline">
                      {isCollapsed ? "Mostrar fotos" : "Plegar sección"}
                    </span>
                    <div className="p-1.5 rounded-lg bg-white/5 text-gray-300">
                      {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
                    </div>
                  </div>
                </button>

                {/* CONTENIDO DESPLEGABLE: GRILLA DE FOTOS DE LA SECCIÓN */}
                {!isCollapsed && (
                  <div className="p-4 sm:p-6 bg-[#0e0e18]">
                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-3.5 sm:gap-4">
                      {secImages.map((img) => (
                        <div
                          key={img.id}
                          className="group relative rounded-2xl bg-[#141422] border border-white/10 flex flex-col hover:border-[#ffd025] transition-all duration-200 hover:shadow-xl hover:shadow-[#ffd025]/10"
                        >
                          {/* Thumbnail Container */}
                          <div
                            onClick={() => setPreviewImage(img)}
                            className="relative aspect-square w-full bg-[#0a0a10] cursor-pointer overflow-hidden rounded-t-2xl flex items-center justify-center"
                          >
                            <img
                              src={img.url}
                              alt={img.name}
                              loading="lazy"
                              className="w-full h-full object-contain p-2 transition-transform duration-300 group-hover:scale-105"
                              onError={(e) => {
                                (e.currentTarget as HTMLImageElement).src =
                                  "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                              }}
                            />

                            {/* Badge categoría identificada */}
                            <div className="absolute top-2 left-2 z-10 pointer-events-none">
                              <span className="px-1.5 py-0.5 rounded-md bg-black/80 border border-white/20 text-gray-200 text-[8px] sm:text-[9px] font-bold uppercase tracking-wider shadow">
                                {secName}
                              </span>
                            </div>

                            {/* Hover overlay actions */}
                            <div className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-between p-2">
                              <div className="w-full flex items-center justify-between">
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleCopyUrl(img.url, img.id);
                                  }}
                                  title="Copiar URL"
                                  className="p-1.5 rounded-lg bg-black/80 hover:bg-[#ffd025] text-white hover:text-black transition cursor-pointer"
                                >
                                  {copiedId === img.id ? <Check size={13} /> : <Copy size={13} />}
                                </button>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleDeleteImage(img.id);
                                  }}
                                  title="Eliminar foto"
                                  className="p-1.5 rounded-lg bg-red-600/80 hover:bg-red-600 text-white transition cursor-pointer"
                                >
                                  <Trash2 size={13} />
                                </button>
                              </div>

                              <span className="text-[10px] text-[#ffd025] font-bold uppercase tracking-wider bg-black/70 px-2 py-0.5 rounded-full">
                                🔍 Ver grande
                              </span>
                            </div>
                          </div>

                          {/* Info footer & Quick Category Switcher & Product Assigner */}
                          <div className="p-2.5 border-t border-white/5 space-y-1.5 flex-1 flex flex-col justify-between">
                            <div className="space-y-1.5">
                              <p className="text-[10.5px] sm:text-[11px] font-bold text-white truncate" title={img.name}>
                                {img.name}
                              </p>

                              {/* Cambiar categoría manualmente */}
                              <div className="flex items-center gap-1">
                                <select
                                  value={img.category || secName}
                                  onChange={(e) => handleUpdateCategory(img.id, e.target.value)}
                                  className="w-full bg-[#0c0c14] border border-white/10 rounded-lg px-1.5 py-1 text-[9.5px] font-semibold text-gray-300 focus:border-[#ffd025] focus:outline-none cursor-pointer"
                                  title="Cambiar sección de esta imagen"
                                >
                                  {ORDERED_SECTIONS.map((s) => (
                                    <option key={s} value={s} className="bg-[#141422] text-white">
                                      {s}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </div>

                            {/* Badges de asignación y botón para abrir buscador predictivo espacioso */}
                            <ImageAssignedBadges
                              imageUrl={img.url}
                              products={products}
                              onOpenAssignModal={() => setAssigningImage(img)}
                              onUnassign={handleUnassignImageFromProduct}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Lightbox Preview Modal */}
      {previewImage && (
        <div
          onClick={() => setPreviewImage(null)}
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-md animate-fade-in"
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="relative max-w-2xl w-full bg-[#141422] border-2 border-[#ffd025]/50 rounded-3xl overflow-hidden p-4 sm:p-6 shadow-2xl flex flex-col items-center max-h-[90vh] overflow-y-auto"
          >
            <div className="w-full flex items-center justify-between mb-4">
              <div className="flex items-center gap-2 truncate max-w-md">
                <span className="px-2 py-0.5 rounded-md bg-[#ffd025]/20 text-[#ffd025] text-xs font-bold uppercase">
                  {previewImage.category || "General"}
                </span>
                <h4 className="text-sm sm:text-base font-bold text-white truncate">
                  {previewImage.name}
                </h4>
              </div>
              <button
                onClick={() => setPreviewImage(null)}
                className="p-1.5 rounded-xl bg-white/10 hover:bg-white/20 text-white transition cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="w-full max-h-[48vh] bg-black/50 rounded-2xl overflow-hidden flex items-center justify-center p-2 mb-4">
              <img
                src={previewImage.url}
                alt={previewImage.name}
                className="max-h-[45vh] max-w-full object-contain"
              />
            </div>

            {/* Asignar a producto desde el preview modal */}
            <div className="w-full bg-[#0e0e18] p-4 rounded-2xl border border-white/10 mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-left">
              <div>
                <div className="flex items-center gap-1.5 text-xs font-bold text-[#ffd025] uppercase tracking-wider">
                  <Sparkles size={13} /> Asignar esta foto a un producto
                </div>
                <p className="text-[11px] text-gray-400 mt-0.5">
                  Abre la ventana de búsqueda predictiva para vincular esta imagen a productos del catálogo.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setAssigningImage(previewImage)}
                className="w-full sm:w-auto px-5 py-2.5 rounded-xl bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] hover:scale-105 active:scale-95 text-black font-black uppercase text-xs tracking-wider transition-all flex items-center justify-center gap-2 shadow-lg shadow-[#ffd025]/20 cursor-pointer shrink-0"
              >
                <Search size={14} strokeWidth={2.5} />
                <span>Asignar a Producto</span>
              </button>
            </div>

            <div className="w-full flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-white/10">
              <span className="text-xs text-gray-400 font-mono truncate max-w-xs">
                {previewImage.url}
              </span>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleCopyUrl(previewImage.url, previewImage.id)}
                  className="px-4 py-2 bg-[#ffd025] hover:bg-[#ffe066] text-black font-bold text-xs rounded-xl transition flex items-center gap-1.5 cursor-pointer"
                >
                  {copiedId === previewImage.id ? <Check size={14} /> : <Copy size={14} />}
                  <span>{copiedId === previewImage.id ? "¡Copiada!" : "Copiar Enlace"}</span>
                </button>
                <button
                  onClick={() => handleDeleteImage(previewImage.id)}
                  className="px-3.5 py-2 bg-red-500/10 hover:bg-red-500/20 text-red-400 border border-red-500/20 font-bold text-xs rounded-xl transition cursor-pointer"
                >
                  Eliminar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal Dedicado y Espacioso de Asignación a Producto */}
      {assigningImage && (
        <AssignProductModal
          image={assigningImage}
          products={products}
          onClose={() => setAssigningImage(null)}
          onAssign={handleAssignImageToProduct}
          onUnassign={handleUnassignImageFromProduct}
        />
      )}

      {/* Floating Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-[#161628]/95 border-2 border-[#ffd025] text-white px-5 py-3 rounded-2xl shadow-2xl flex items-center gap-2.5 animate-bounce-in text-xs font-bold backdrop-blur-md">
          <CheckCircle2 size={16} className="text-[#ffd025] shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}
    </div>
  );
};
