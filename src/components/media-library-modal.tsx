import React, { useState, useEffect, useRef, useMemo } from "react";
import { createPortal } from "react-dom";
import {
  Upload,
  Image as ImageIcon,
  Trash2,
  Copy,
  Check,
  X,
  Search,
  RefreshCw,
  FolderOpen,
  Sparkles,
  Eye,
  EyeOff,
  ChevronDown,
  ChevronUp,
  Layers,
  Bot,
  Flame,
  Wine,
  Beer,
  GlassWater,
  Cigarette,
  Cookie,
  Candy,
  Package,
  Tag,
} from "lucide-react";
import { uploadImagesBatch } from "../lib/image-batch-uploader";
import { CATEGORY_ICONS, ORDERED_SECTIONS } from "./media-admin-panel";

export interface MediaItem {
  id: string;
  name: string;
  url: string;
  createdAt: string;
  size?: number;
  category?: string;
  aiDetectedTitle?: string;
}

interface MediaLibraryModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectImage?: (url: string) => void;
  onMediaDeleted?: () => void;
  assignedImages?: string[];
  currentImage?: string;
  title?: string;
}

export const MediaLibraryModal: React.FC<MediaLibraryModalProps> = ({
  isOpen,
  onClose,
  onSelectImage,
  onMediaDeleted,
  assignedImages = [],
  currentImage = "",
  title = "Archivo y Galería de Imágenes",
}) => {
  const [images, setImages] = useState<MediaItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState("");
  const [uploadPercent, setUploadPercent] = useState(0);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedSectionFilter, setSelectedSectionFilter] = useState<string>("all");
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [hideAssigned, setHideAssigned] = useState(true);
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
    if (isOpen) {
      fetchImages();
      // Lock scroll on background body
      const originalOverflow = document.body.style.overflow;
      document.body.style.overflow = "hidden";

      const handleKeyDown = (e: KeyboardEvent) => {
        if (e.key === "Escape") {
          onClose();
        }
      };
      window.addEventListener("keydown", handleKeyDown);

      return () => {
        document.body.style.overflow = originalOverflow;
        window.removeEventListener("keydown", handleKeyDown);
      };
    }
  }, [isOpen]);

  if (!isOpen) return null;

  // Process batch of files with concurrency, compression, and auto AI classification
  const handleBatchUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;

    const filesArray = Array.from(fileList);
    setUploading(true);
    setUploadPercent(0);
    setUploadProgress(`Iniciando carga y análisis de ${filesArray.length} imágenes...`);

    try {
      const { successCount, failedCount } = await uploadImagesBatch(
        filesArray,
        (p) => {
          setUploadPercent(p.percent);
          setUploadProgress(
            `Procesando y clasificando: ${p.current} de ${p.total} (${p.percent}%) — ${p.fileName}`
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

  const handleDeleteImage = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm("¿Deseas eliminar permanentemente esta imagen del archivo?")) return;

    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/${id}`, {
        method: "DELETE",
      });
      if (res.ok) {
        setImages((prev) => prev.filter((img) => img.id !== id));
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
        onMediaDeleted?.();
        alert("Se han eliminado todas las imágenes del archivo de medios y de los productos con éxito.");
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

  const handleCopyUrl = (url: string, id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    navigator.clipboard.writeText(url);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const handleUpdateCategory = async (id: string, newCategory: string, e?: React.ChangeEvent<HTMLSelectElement>) => {
    e?.stopPropagation();
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

  const isImageAssigned = (img: MediaItem) => {
    if (!assignedImages || assignedImages.length === 0) return false;
    return assignedImages.some((assigned) => {
      if (!assigned) return false;
      const cleanAssigned = String(assigned).trim().toLowerCase();
      const cleanUrl = String(img.url).trim().toLowerCase();
      const cleanId = String(img.id).trim().toLowerCase();
      const cleanName = String(img.name).trim().toLowerCase();
      return (
        cleanAssigned === cleanUrl ||
        cleanAssigned === cleanId ||
        cleanAssigned.includes(cleanId) ||
        cleanAssigned === cleanName ||
        cleanAssigned.endsWith("/" + cleanId)
      );
    });
  };

  const isCurrentProductImage = (img: MediaItem) => {
    if (!currentImage) return false;
    const cleanCurrent = String(currentImage).trim().toLowerCase();
    const cleanUrl = String(img.url).trim().toLowerCase();
    const cleanId = String(img.id).trim().toLowerCase();
    return cleanCurrent === cleanUrl || cleanCurrent === cleanId || cleanCurrent.includes(cleanId);
  };

  const assignedCount = images.filter(isImageAssigned).length;

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

  // Group images by section taking into account search query and assigned filter
  const { groupedImages, sectionCounts, allFilteredCount } = useMemo(() => {
    const counts: Record<string, number> = {};
    ORDERED_SECTIONS.forEach((s) => (counts[s] = 0));

    const q = searchQuery.toLowerCase().trim();

    const filtered = images.filter((img) => {
      // 1. Search match
      if (q) {
        const matchName = img.name.toLowerCase().includes(q);
        const matchCat = (img.category || "").toLowerCase().includes(q);
        const matchTitle = (img.aiDetectedTitle || "").toLowerCase().includes(q);
        if (!matchName && !matchCat && !matchTitle) return false;
      }

      // 2. Assigned filter
      if (hideAssigned && isImageAssigned(img)) {
        // If this image is the one currently assigned to this product, keep it visible
        if (isCurrentProductImage(img)) return true;
        return false;
      }

      return true;
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
  }, [images, searchQuery, hideAssigned, assignedImages, currentImage]);

  const modalNode = (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        zIndex: 999999,
      }}
      className="w-full h-full max-w-full max-h-full overflow-hidden flex items-center justify-center p-2 sm:p-4 md:p-6 bg-black/85 backdrop-blur-md"
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="relative w-full max-w-[98vw] md:max-w-5xl lg:max-w-6xl h-[92vh] max-h-[900px] bg-[#12121e] border-2 border-[#ffd025]/60 rounded-3xl shadow-2xl flex flex-col overflow-hidden text-left ring-4 ring-black/70 animate-scale-in"
      >
        {/* Modal Header */}
        <div className="px-4 sm:px-6 py-3 sm:py-4 border-b border-white/10 flex items-center justify-between bg-[#171726] shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-2xl bg-[#ffd025]/15 border border-[#ffd025]/30 flex items-center justify-center text-[#ffd025] shrink-0">
              <FolderOpen size={20} />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <h3 className="text-sm sm:text-base md:text-lg font-black text-white uppercase tracking-wider truncate">
                  {title}
                </h3>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#ffd025]/20 text-[#ffd025] font-mono font-bold shrink-0">
                  {images.length} fotos guardadas
                </span>
              </div>
              <p className="text-[11px] sm:text-xs text-gray-400 truncate">
                {onSelectImage
                  ? "Organizadas por secciones con IA. Toca una foto para seleccionarla."
                  : "Organizadas por secciones con IA para rápida búsqueda."}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            {images.length > 0 && (
              <button
                type="button"
                onClick={handleClassifyAI}
                disabled={loading || classifyingAI || uploading}
                title="Clasificar y organizar fotos con IA"
                className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-purple-600/20 hover:bg-purple-600/30 text-purple-300 border border-purple-500/30 text-xs font-bold transition cursor-pointer"
              >
                <Bot size={14} className={classifyingAI ? "animate-bounce text-[#ffd025]" : ""} />
                <span>{classifyingAI ? "Analizando..." : "Re-organizar con IA"}</span>
              </button>
            )}

            <button
              type="button"
              onClick={fetchImages}
              disabled={loading}
              title="Recargar archivo"
              className="p-2 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 transition-colors cursor-pointer"
            >
              <RefreshCw size={16} className={loading ? "animate-spin text-[#ffd025]" : ""} />
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl bg-red-500/10 hover:bg-red-500/20 text-red-400 transition-colors cursor-pointer"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* Action Controls & Batch Uploader */}
        <div className="p-3 sm:p-4 border-b border-white/10 bg-[#141422] flex flex-col sm:flex-row gap-2.5 sm:gap-3 items-center justify-between shrink-0">
          {/* Search Box & Assigned Filter Toggle */}
          <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
            <div className="relative w-full sm:w-60">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                placeholder="Buscar por nombre o tipo..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-2 bg-[#0c0c14] border border-white/10 rounded-xl text-white text-xs focus:border-[#ffd025] focus:outline-none"
              />
            </div>

            {/* BOTÓN FILTRO DE ASIGNADAS: Ocultar o ver asignadas */}
            {assignedCount > 0 && onSelectImage && (
              <button
                type="button"
                onClick={() => setHideAssigned(!hideAssigned)}
                className={`px-3 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all border cursor-pointer shrink-0 ${
                  hideAssigned
                    ? "bg-[#ffd025]/15 text-[#ffd025] border-[#ffd025]/30 hover:bg-[#ffd025]/25"
                    : "bg-white/5 text-gray-400 border-white/10 hover:text-white"
                }`}
                title={hideAssigned ? "Mostrando solo fotos libres (sin asignar a otros productos)" : "Mostrando todas las fotos"}
              >
                {hideAssigned ? <EyeOff size={14} /> : <Eye size={14} />}
                <span>{hideAssigned ? `Ocultando ${assignedCount} ya asignadas` : `Ver todas (${assignedCount} asignadas)`}</span>
              </button>
            )}
          </div>

          {/* Big Batch Upload Button & Delete All */}
          <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto justify-end">
            <button
              type="button"
              onClick={() => toggleAllSections(false)}
              className="px-2.5 py-2 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 hover:text-white text-[11px] font-bold transition border border-white/10 cursor-pointer hidden md:block"
            >
              Expandir todo
            </button>
            <button
              type="button"
              onClick={() => toggleAllSections(true)}
              className="px-2.5 py-2 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 hover:text-white text-[11px] font-bold transition border border-white/10 cursor-pointer hidden md:block"
            >
              Plegar todo
            </button>

            {images.length > 0 && (
              <button
                type="button"
                onClick={handleDeleteAllImages}
                disabled={loading || uploading}
                className="px-3 py-2 rounded-xl bg-red-600/20 hover:bg-red-600/30 text-red-400 hover:text-red-300 border border-red-500/30 transition-all font-bold text-xs flex items-center gap-1.5 cursor-pointer shrink-0"
                title="Eliminar todas las imágenes del archivo de medios"
              >
                <Trash2 size={14} />
                <span>Borrar Todas ({images.length})</span>
              </button>
            )}

            <label
              className={`cursor-pointer px-4 sm:px-5 py-2 rounded-xl font-black uppercase text-xs tracking-wider transition-all flex items-center justify-center gap-2 shadow-lg shadow-[#ffd025]/20 shrink-0 ${
                uploading
                  ? "bg-gray-700 text-gray-400 cursor-not-allowed"
                  : "bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] hover:scale-[1.02] active:scale-[0.98]"
              }`}
            >
              <Upload size={15} strokeWidth={2.5} />
              <span>{uploading ? "Subiendo lote..." : "Subir Fotos"}</span>
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

        {/* Upload Progress Banner & Bar */}
        {uploading && (
          <div className="px-5 py-2.5 bg-[#ffd025]/15 border-b border-[#ffd025]/30 space-y-1.5 animate-fade-in shadow-inner shrink-0">
            <div className="flex items-center justify-between gap-3 text-xs font-bold text-[#ffd025]">
              <span className="flex items-center gap-2 truncate">
                <RefreshCw size={14} className="animate-spin text-[#ffd025] shrink-0" />
                <span className="truncate">{uploadProgress}</span>
              </span>
              <span className="font-mono text-xs font-black shrink-0">{uploadPercent}%</span>
            </div>
            <div className="w-full h-1.5 rounded-full bg-black/50 overflow-hidden border border-white/10">
              <div
                className="h-full bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] transition-all duration-300 rounded-full"
                style={{ width: `${uploadPercent}%` }}
              />
            </div>
          </div>
        )}

        {/* SELECTOR RÁPIDO DE SECCIONES (BARRA HORIZONTAL CON SCROLL) */}
        {images.length > 0 && (
          <div className="px-4 py-2 border-b border-white/10 bg-[#10101c] flex items-center gap-1.5 overflow-x-auto scrollbar-none shrink-0">
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider mr-1 shrink-0 flex items-center gap-1">
              <Layers size={12} className="text-[#ffd025]" /> Secciones:
            </span>

            <button
              type="button"
              onClick={() => setSelectedSectionFilter("all")}
              className={`px-2.5 py-1 rounded-lg text-[11px] font-bold uppercase tracking-wider whitespace-nowrap transition flex items-center gap-1.5 cursor-pointer border shrink-0 ${
                selectedSectionFilter === "all"
                  ? "bg-[#ffd025] text-black border-[#ffd025] font-black shadow-sm"
                  : "bg-white/5 text-gray-400 hover:text-white border-white/10"
              }`}
            >
              <span>Todas ({allFilteredCount})</span>
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
                    setCollapsedSections((prev) => ({ ...prev, [secName]: false }));
                  }}
                  className={`px-2.5 py-1 rounded-lg text-[11px] font-bold whitespace-nowrap transition flex items-center gap-1.5 cursor-pointer border shrink-0 ${
                    isSelected
                      ? "bg-[#ffd025] text-black border-[#ffd025] font-black shadow-sm"
                      : "bg-white/5 text-gray-300 hover:text-white border-white/10 hover:border-white/20"
                  }`}
                >
                  {CATEGORY_ICONS[secName] || <Tag size={12} />}
                  <span>{secName}</span>
                  <span
                    className={`text-[9px] px-1.5 py-0.2 rounded-full font-mono font-bold ${
                      isSelected ? "bg-black/20 text-black" : "bg-white/10 text-gray-300"
                    }`}
                  >
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        )}

        {/* Scrollable Content Area: Sections with Collapsible Accordions */}
        <div className="flex-1 overflow-y-auto p-3 sm:p-5 bg-[#0e0e18] custom-admin-scrollbar space-y-4">
          {loading && images.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-gray-500 py-16">
              <RefreshCw size={32} className="animate-spin text-[#ffd025] mb-3" />
              <p className="text-sm font-semibold">Cargando archivo de imágenes...</p>
            </div>
          ) : allFilteredCount === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-gray-500 py-16 text-center">
              <div className="w-16 h-16 rounded-3xl bg-white/5 border border-white/10 flex items-center justify-center mb-4 text-[#ffd025]/60">
                <ImageIcon size={32} />
              </div>
              <h4 className="text-base font-bold text-white mb-1">
                {searchQuery
                  ? "No se encontraron fotos con esa búsqueda"
                  : hideAssigned && assignedCount > 0
                  ? "Todas las fotos ya están asignadas a productos"
                  : "El archivo está vacío"}
              </h4>
              <p className="text-xs text-gray-400 max-w-sm mb-5">
                {searchQuery
                  ? "Prueba buscando con otro término."
                  : hideAssigned && assignedCount > 0
                  ? "Haz clic en 'Ver todas' para ver las fotos ya asignadas o sube nuevas fotos."
                  : "Usa el botón superior para seleccionar varias fotos desde tu dispositivo y subirlas todas juntas."}
              </p>
              {hideAssigned && assignedCount > 0 ? (
                <button
                  type="button"
                  onClick={() => setHideAssigned(false)}
                  className="px-5 py-2.5 rounded-xl bg-[#ffd025] text-black font-black uppercase text-xs tracking-wider hover:bg-[#ffe066] transition cursor-pointer"
                >
                  Ver Todas ({assignedCount} Asignadas)
                </button>
              ) : (
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="px-5 py-2.5 rounded-xl bg-[#ffd025] text-black font-black uppercase text-xs tracking-wider hover:bg-[#ffe066] transition cursor-pointer"
                >
                  Subir Fotos Ahora
                </button>
              )}
            </div>
          ) : (
            <div className="space-y-4">
              {ORDERED_SECTIONS.map((secName) => {
                // If a specific section is filtered, only show that section
                if (selectedSectionFilter !== "all" && selectedSectionFilter !== secName) {
                  return null;
                }

                const sectionImages = groupedImages[secName] || [];
                if (sectionImages.length === 0) return null;

                const isCollapsed = collapsedSections[secName] || false;

                return (
                  <div
                    key={secName}
                    className="rounded-2xl bg-[#131320] border border-white/10 overflow-hidden shadow-lg transition-all"
                  >
                    {/* Collapsible Section Header */}
                    <div
                      onClick={() => toggleSection(secName)}
                      className="w-full px-4 py-3 bg-[#171726] hover:bg-[#1c1c2e] border-b border-white/5 flex items-center justify-between cursor-pointer select-none transition-colors"
                    >
                      <div className="flex items-center gap-2.5">
                        <div className="p-1.5 rounded-xl bg-white/5 border border-white/10 text-[#ffd025]">
                          {CATEGORY_ICONS[secName] || <Tag size={16} />}
                        </div>
                        <span className="text-sm font-black text-white uppercase tracking-wider">
                          {secName}
                        </span>
                        <span className="text-xs px-2 py-0.5 rounded-full bg-[#ffd025]/20 text-[#ffd025] font-mono font-bold">
                          {sectionImages.length} fotos
                        </span>
                      </div>

                      <div className="flex items-center gap-2 text-gray-400">
                        <span className="text-[11px] hidden sm:inline text-gray-500 font-semibold">
                          {isCollapsed ? "Clic para desplegar" : "Clic para plegar"}
                        </span>
                        {isCollapsed ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
                      </div>
                    </div>

                    {/* Section Grid of Images (Collapsible Content) */}
                    {!isCollapsed && (
                      <div className="p-3 sm:p-4 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-2.5 sm:gap-3.5 animate-fade-in">
                        {sectionImages.map((img) => {
                          const assigned = isImageAssigned(img);
                          const isCurrent = isCurrentProductImage(img);

                          return (
                            <div
                              key={img.id}
                              onClick={() => {
                                if (onSelectImage) {
                                  onSelectImage(img.url);
                                  onClose();
                                }
                              }}
                              className={`group relative rounded-2xl bg-[#161624] border border-white/10 overflow-hidden flex flex-col transition-all duration-200 hover:border-[#ffd025] hover:shadow-xl hover:shadow-[#ffd025]/10 ${
                                onSelectImage ? "cursor-pointer hover:scale-[1.02]" : ""
                              } ${isCurrent ? "border-2 border-[#ffd025] ring-2 ring-[#ffd025]/30 bg-[#ffd025]/5" : ""}`}
                            >
                              {/* Image Aspect Box */}
                              <div className="relative aspect-square w-full bg-[#09090f] overflow-hidden flex items-center justify-center">
                                {isCurrent ? (
                                  <div className="absolute top-2 left-2 z-10 pointer-events-none">
                                    <span className="px-1.5 py-0.5 rounded-md bg-[#ffd025] text-black text-[8.5px] font-black uppercase tracking-wider shadow-md">
                                      Foto Actual
                                    </span>
                                  </div>
                                ) : assigned ? (
                                  <div className="absolute top-2 left-2 z-10 pointer-events-none">
                                    <span className="px-1.5 py-0.5 rounded-md bg-black/85 border border-[#ffd025]/50 text-[#ffd025] text-[8.5px] font-black uppercase tracking-wider shadow-md">
                                      Ya Asignada
                                    </span>
                                  </div>
                                ) : null}

                                <img
                                  src={img.url}
                                  alt={img.name}
                                  loading="lazy"
                                  className="w-full h-full object-contain p-1.5 transition-transform duration-300 group-hover:scale-105"
                                  onError={(e) => {
                                    (e.currentTarget as HTMLImageElement).src =
                                      "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                  }}
                                />

                                {/* Overlay on hover */}
                                <div className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-between p-2">
                                  <div className="w-full flex items-center justify-between">
                                    <button
                                      type="button"
                                      onClick={(e) => handleCopyUrl(img.url, img.id, e)}
                                      title="Copiar URL"
                                      className="p-1.5 rounded-lg bg-black/70 hover:bg-[#ffd025] text-white hover:text-black transition cursor-pointer"
                                    >
                                      {copiedId === img.id ? <Check size={13} /> : <Copy size={13} />}
                                    </button>
                                    <button
                                      type="button"
                                      onClick={(e) => handleDeleteImage(img.id, e)}
                                      title="Eliminar de la galería"
                                      className="p-1.5 rounded-lg bg-black/70 hover:bg-red-600 text-white transition cursor-pointer"
                                    >
                                      <Trash2 size={13} />
                                    </button>
                                  </div>

                                  {onSelectImage && (
                                    <span className="px-2.5 py-1 rounded-lg bg-[#ffd025] text-black font-black text-[10px] uppercase tracking-wider shadow-lg">
                                      Seleccionar
                                    </span>
                                  )}

                                  <div className="text-[9px] text-gray-300 truncate w-full text-center">
                                    {img.size ? `${(img.size / 1024).toFixed(0)} KB` : "Imagen"}
                                  </div>
                                </div>
                              </div>

                              {/* Card Meta & Title */}
                              <div className="p-2 border-t border-white/5 bg-[#12121e] flex flex-col justify-between space-y-1" onClick={(e) => e.stopPropagation()}>
                                <p className="text-[11px] font-bold text-gray-200 truncate" title={img.name}>
                                  {img.name}
                                </p>
                                {img.aiDetectedTitle && (
                                  <p className="text-[9.5px] text-[#ffd025] truncate font-medium flex items-center gap-1">
                                    <Sparkles size={10} className="shrink-0" />
                                    <span className="truncate">{img.aiDetectedTitle}</span>
                                  </p>
                                )}
                                <select
                                  value={img.category || secName}
                                  onChange={(e) => handleUpdateCategory(img.id, e.target.value, e)}
                                  onClick={(e) => e.stopPropagation()}
                                  className="w-full bg-[#0c0c14] border border-white/10 rounded-lg px-1.5 py-0.5 text-[9px] font-semibold text-gray-300 focus:border-[#ffd025] focus:outline-none cursor-pointer"
                                  title="Mover foto a otra sección"
                                >
                                  {ORDERED_SECTIONS.map((s) => (
                                    <option key={s} value={s} className="bg-[#141422] text-white">
                                      {s}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-5 py-3 border-t border-white/10 bg-[#141422] flex items-center justify-between text-xs text-gray-400 shrink-0">
          <div className="flex items-center gap-3">
            <span className="font-semibold text-white">
              Mostrando {allFilteredCount} de {images.length} fotos
            </span>
            {assignedCount > 0 && (
              <span className="text-[11px] text-[#ffd025] font-medium hidden sm:inline">
                ({assignedCount} asignadas a productos)
              </span>
            )}
          </div>

          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2 rounded-xl bg-white/10 hover:bg-white/15 text-white font-bold transition cursor-pointer text-xs uppercase tracking-wider"
          >
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );

  return createPortal(modalNode, document.body);
};
