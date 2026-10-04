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

interface MediaAdminPanelProps {
  onMediaDeleted?: () => void;
}

export const CATEGORY_ICONS: Record<string, React.ReactNode> = {
  "Promociones & Packs": <Flame size={16} className="text-amber-400" />,
  "Piscos": <Wine size={16} className="text-[#ffd025]" />,
  "Cervezas": <Beer size={16} className="text-yellow-400" />,
  "Vinos & Espumantes": <Wine size={16} className="text-purple-400" />,
  "Destilados & Licores": <GlassWater size={16} className="text-cyan-400" />,
  "Bebidas & Energéticas": <GlassWater size={16} className="text-emerald-400" />,
  "Cigarros & Vapes": <Cigarette size={16} className="text-orange-400" />,
  "Snacks & Piqueos": <Cookie size={16} className="text-amber-500" />,
  "Dulces & Chocolates": <Candy size={16} className="text-pink-400" />,
  "Hielo & Abarrotes": <Package size={16} className="text-blue-400" />,
  "Otros / General": <FolderOpen size={16} className="text-gray-400" />,
};

export const ORDERED_SECTIONS = [
  "Promociones & Packs",
  "Piscos",
  "Cervezas",
  "Vinos & Espumantes",
  "Destilados & Licores",
  "Bebidas & Energéticas",
  "Cigarros & Vapes",
  "Snacks & Piqueos",
  "Dulces & Chocolates",
  "Hielo & Abarrotes",
  "Otros / General",
];

export const MediaAdminPanel: React.FC<MediaAdminPanelProps> = ({ onMediaDeleted }) => {
  const [images, setImages] = useState<MediaItem[]>([]);
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
  const fileInputRef = useRef<HTMLInputElement>(null);

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

    const q = searchQuery.toLowerCase().trim();

    const filtered = images.filter((img) => {
      if (!q) return true;
      const matchName = img.name.toLowerCase().includes(q);
      const matchCat = (img.category || "").toLowerCase().includes(q);
      const matchTitle = (img.aiDetectedTitle || "").toLowerCase().includes(q);
      return matchName || matchCat || matchTitle;
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
  }, [images, searchQuery]);

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
                          className="group relative rounded-2xl bg-[#141422] border border-white/10 overflow-hidden flex flex-col hover:border-[#ffd025] transition-all duration-200 hover:shadow-xl hover:shadow-[#ffd025]/10"
                        >
                          {/* Thumbnail Container */}
                          <div
                            onClick={() => setPreviewImage(img)}
                            className="relative aspect-square w-full bg-[#0a0a10] cursor-pointer overflow-hidden flex items-center justify-center"
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

                          {/* Info footer & Quick Category Switcher */}
                          <div className="p-2.5 border-t border-white/5 space-y-1.5">
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
            className="relative max-w-2xl w-full bg-[#141422] border-2 border-[#ffd025]/50 rounded-3xl overflow-hidden p-4 sm:p-6 shadow-2xl flex flex-col items-center"
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

            <div className="w-full max-h-[60vh] bg-black/50 rounded-2xl overflow-hidden flex items-center justify-center p-2 mb-4">
              <img
                src={previewImage.url}
                alt={previewImage.name}
                className="max-h-[55vh] max-w-full object-contain"
              />
            </div>

            <div className="w-full flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-white/10">
              <span className="text-xs text-gray-400 font-mono">
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
    </div>
  );
};
