import React, { useState, useEffect, useRef } from "react";
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
  ExternalLink,
  Eye,
  EyeOff,
} from "lucide-react";
import { uploadImagesBatch } from "../lib/image-batch-uploader";

export interface MediaItem {
  id: string;
  name: string;
  url: string;
  createdAt: string;
  size?: number;
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
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [hideAssigned, setHideAssigned] = useState(true);
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

  // Process batch of files with concurrency and compression
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

  const assignedCount = images.filter(isImageAssigned).length;

  const filteredImages = images.filter((img) => {
    const matchesSearch = img.name.toLowerCase().includes(searchQuery.toLowerCase());
    if (!matchesSearch) return false;
    if (hideAssigned && isImageAssigned(img)) {
      // Si esta imagen coincide con la que ya tiene asignada este producto en particular, la mostramos para permitir conservarla
      if (currentImage) {
        const cleanCurrent = String(currentImage).trim().toLowerCase();
        const cleanUrl = String(img.url).trim().toLowerCase();
        const cleanId = String(img.id).trim().toLowerCase();
        if (cleanCurrent === cleanUrl || cleanCurrent === cleanId || cleanCurrent.includes(cleanId)) {
          return true;
        }
      }
      return false;
    }
    return true;
  });

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
      className="w-full h-full max-w-full max-h-full overflow-hidden flex items-center justify-center p-3 sm:p-6 bg-black/85 backdrop-blur-md"
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="relative w-full max-w-[95vw] md:max-w-4xl lg:max-w-5xl h-[88vh] max-h-[850px] bg-[#12121e] border-2 border-[#ffd025]/60 rounded-3xl shadow-2xl flex flex-col overflow-hidden text-left ring-4 ring-black/70 animate-scale-in"
      >
        {/* Modal Header */}
        <div className="px-5 py-4 border-b border-white/10 flex items-center justify-between bg-[#171726] shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-[#ffd025]/15 border border-[#ffd025]/30 flex items-center justify-center text-[#ffd025]">
              <FolderOpen size={20} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base sm:text-lg font-black text-white uppercase tracking-wider">
                  {title}
                </h3>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#ffd025]/20 text-[#ffd025] font-mono font-bold">
                  {images.length} fotos guardadas
                </span>
              </div>
              <p className="text-xs text-gray-400">
                {onSelectImage
                  ? "Haz clic sobre cualquier imagen para seleccionarla para este producto."
                  : "Sube grupos de fotos de golpe para tenerlas siempre disponibles."}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
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
        <div className="p-4 sm:p-5 border-b border-white/10 bg-[#141422] flex flex-col sm:flex-row gap-3 items-center justify-between shrink-0">
          {/* Search Box & Assigned Filter Toggle */}
          <div className="flex flex-wrap items-center gap-2.5 w-full sm:w-auto">
            <div className="relative w-full sm:w-64">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                placeholder="Buscar por nombre..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-2 bg-[#0c0c14] border border-white/10 rounded-xl text-white text-xs focus:border-[#ffd025] focus:outline-none"
              />
            </div>

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
          <div className="flex flex-wrap items-center gap-2.5 w-full sm:w-auto justify-end">
            {images.length > 0 && (
              <button
                type="button"
                onClick={handleDeleteAllImages}
                disabled={loading || uploading}
                className="px-3.5 py-2.5 rounded-xl bg-red-600/20 hover:bg-red-600/30 text-red-400 hover:text-red-300 border border-red-500/30 transition-all font-bold text-xs flex items-center gap-1.5 cursor-pointer shrink-0"
                title="Eliminar todas las imágenes del archivo de medios"
              >
                <Trash2 size={15} />
                <span>Borrar Todas ({images.length})</span>
              </button>
            )}

            <label
              className={`cursor-pointer px-5 py-2.5 rounded-xl font-black uppercase text-xs tracking-wider transition-all flex items-center justify-center gap-2 shadow-lg shadow-[#ffd025]/20 shrink-0 ${
                uploading
                  ? "bg-gray-700 text-gray-400 cursor-not-allowed"
                  : "bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] hover:scale-[1.02] active:scale-[0.98]"
              }`}
            >
              <Upload size={16} strokeWidth={2.5} />
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

        {/* Upload Progress Banner & Bar */}
        {uploading && (
          <div className="px-5 py-3 bg-[#ffd025]/15 border-b border-[#ffd025]/30 space-y-2 animate-fade-in shadow-inner shrink-0">
            <div className="flex items-center justify-between gap-3 text-xs font-bold text-[#ffd025]">
              <span className="flex items-center gap-2 truncate">
                <RefreshCw size={15} className="animate-spin text-[#ffd025] shrink-0" />
                <span className="truncate">{uploadProgress}</span>
              </span>
              <span className="font-mono text-xs font-black shrink-0">{uploadPercent}%</span>
            </div>
            <div className="w-full h-2 rounded-full bg-black/50 overflow-hidden border border-white/10">
              <div
                className="h-full bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] transition-all duration-300 rounded-full"
                style={{ width: `${uploadPercent}%` }}
              />
            </div>
          </div>
        )}

        {/* Grid of Images */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-[#0e0e18] custom-admin-scrollbar">
          {loading && images.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-gray-500 py-16">
              <RefreshCw size={32} className="animate-spin text-[#ffd025] mb-3" />
              <p className="text-sm font-semibold">Cargando archivo de imágenes...</p>
            </div>
          ) : filteredImages.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-gray-500 py-16 text-center">
              <div className="w-16 h-16 rounded-3xl bg-white/5 border border-white/10 flex items-center justify-center mb-4 text-[#ffd025]/60">
                <ImageIcon size={32} />
              </div>
              <h4 className="text-base font-bold text-white mb-1">
                {searchQuery ? "No se encontraron fotos con esa búsqueda" : "El archivo está vacío"}
              </h4>
              <p className="text-xs text-gray-400 max-w-sm mb-5">
                {searchQuery
                  ? "Prueba buscando con otro término."
                  : "Usa el botón superior para seleccionar varias fotos desde tu computadora o celular y subirlas todas juntas."}
              </p>
              {!searchQuery && (
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="px-6 py-2.5 rounded-xl bg-[#ffd025] text-black font-black uppercase text-xs tracking-wider hover:bg-[#ffe066] transition cursor-pointer"
                >
                  Seleccionar Fotos Ahora
                </button>
              )}
            </div>
          ) : (
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-3 sm:gap-4">
              {filteredImages.map((img) => (
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
                  }`}
                >
                  {/* Image Aspect Box */}
                  <div className="relative aspect-square w-full bg-[#09090f] overflow-hidden flex items-center justify-center">
                    {isImageAssigned(img) && (
                      <div className="absolute top-2 left-2 z-10 pointer-events-none">
                        <span className="px-1.5 py-0.5 rounded-md bg-black/85 border border-[#ffd025]/50 text-[#ffd025] text-[8.5px] font-black uppercase tracking-wider shadow-md">
                          Ya Asignada
                        </span>
                      </div>
                    )}
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
                          className="p-1.5 rounded-lg bg-black/70 hover:bg-[#ffd025] text-white hover:text-black transition"
                        >
                          {copiedId === img.id ? <Check size={13} /> : <Copy size={13} />}
                        </button>
                        <button
                          type="button"
                          onClick={(e) => handleDeleteImage(img.id, e)}
                          title="Eliminar del archivo"
                          className="p-1.5 rounded-lg bg-red-600/80 hover:bg-red-600 text-white transition"
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>

                      {onSelectImage && (
                        <div className="px-3 py-1 bg-[#ffd025] text-black font-black text-[10px] uppercase rounded-lg shadow-md flex items-center gap-1">
                          <Sparkles size={11} /> Usar esta foto
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Metadata label */}
                  <div className="p-2 border-t border-white/5">
                    <p className="text-[10px] sm:text-[11px] font-bold text-white truncate" title={img.name}>
                      {img.name}
                    </p>
                    <div className="flex items-center justify-between text-[9px] text-gray-500 mt-0.5">
                      <span>{new Date(img.createdAt).toLocaleDateString("es-CL")}</span>
                      {img.size ? <span>{Math.round(img.size / 1024)} KB</span> : null}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-5 py-3 border-t border-white/10 bg-[#171726] flex items-center justify-between shrink-0">
          <span className="text-[11px] text-gray-400">
            {onSelectImage
              ? "💡 Tip: Haz clic en una foto para asignarla automáticamente al producto."
              : "💡 Tip: Puedes subir tantas fotos como quieras y usarlas luego en cualquier producto o banner."}
          </span>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 rounded-xl border border-white/10 text-gray-300 hover:text-white hover:bg-white/5 text-xs font-bold transition-colors cursor-pointer"
          >
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );

  return typeof document !== "undefined" ? createPortal(modalNode, document.body) : null;
};
