import React, { useState, useEffect, useRef } from "react";
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
} from "lucide-react";

export interface MediaItem {
  id: string;
  name: string;
  url: string;
  createdAt: string;
  size?: number;
}

export const MediaAdminPanel: React.FC = () => {
  const [images, setImages] = useState<MediaItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [previewImage, setPreviewImage] = useState<MediaItem | null>(null);
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
    setUploadProgress(`Preparando ${filesArray.length} imágenes...`);

    try {
      const processedFiles: Array<{ name: string; data: string }> = [];

      for (let i = 0; i < filesArray.length; i++) {
        const file = filesArray[i];
        setUploadProgress(`Procesando imagen ${i + 1} de ${filesArray.length}...`);

        const dataUrl = await new Promise<string>((resolve) => {
          const reader = new FileReader();
          reader.onload = (ev) => resolve(ev.target?.result as string);
          reader.readAsDataURL(file);
        });

        processedFiles.push({
          name: file.name,
          data: dataUrl,
        });
      }

      setUploadProgress(`Guardando ${processedFiles.length} imágenes en el servidor...`);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library/upload-batch`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ files: processedFiles }),
      });

      if (res.ok) {
        await fetchImages();
      } else {
        alert("Hubo un error al subir el grupo de imágenes.");
      }
    } catch (err) {
      console.error("Error in batch upload:", err);
      alert("Error al procesar las imágenes seleccionadas.");
    } finally {
      setUploading(false);
      setUploadProgress("");
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
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
      }
    } catch (err) {
      console.error("Error deleting image:", err);
    }
  };

  const handleCopyUrl = (url: string, id: string) => {
    navigator.clipboard.writeText(url);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const filteredImages = images.filter((img) =>
    img.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="space-y-6 max-w-6xl mx-auto animate-fade-in text-left">
      {/* Top Banner Card */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-7 rounded-3xl border border-white/10 flex flex-col md:flex-row md:items-center justify-between gap-5 shadow-xl shadow-black/60">
        <div>
          <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
            <Sparkles size={14} /> Archivo Central de Imágenes
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-white uppercase tracking-wider">
            Galería y Archivo de Medios
          </h2>
          <p className="text-xs text-gray-400 mt-1 max-w-xl">
            Sube múltiples imágenes en grupo o de golpe para tenerlas siempre a mano. Luego podrás seleccionarlas con un solo clic al editar cualquier producto.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={fetchImages}
            disabled={loading}
            className="p-3 rounded-2xl bg-white/5 hover:bg-white/10 text-gray-300 transition-colors border border-white/10 cursor-pointer"
            title="Recargar archivo"
          >
            <RefreshCw size={18} className={loading ? "animate-spin text-[#ffd025]" : ""} />
          </button>

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

      {/* Upload Progress Alert */}
      {uploading && (
        <div className="p-4 rounded-2xl bg-[#ffd025]/15 border border-[#ffd025]/30 flex items-center justify-center gap-3 animate-pulse">
          <RefreshCw size={16} className="animate-spin text-[#ffd025]" />
          <span className="text-xs sm:text-sm font-bold text-[#ffd025]">{uploadProgress}</span>
        </div>
      )}

      {/* Gallery Controls & Search */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-3xl border border-white/10 flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="relative w-full sm:w-80">
          <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            placeholder="Buscar imagen por nombre..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full pl-10 pr-4 py-2.5 bg-[#0e0e16] border border-white/10 rounded-xl text-white text-xs focus:border-[#ffd025] focus:outline-none"
          />
        </div>

        <div className="text-xs text-gray-400 font-mono flex items-center gap-2">
          <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse" />
          <span>{images.length} imágenes en el archivo</span>
        </div>
      </div>

      {/* Grid of Images */}
      {loading && images.length === 0 ? (
        <div className="p-16 rounded-3xl bg-[#13131f]/60 border border-white/10 flex flex-col items-center justify-center text-gray-400">
          <RefreshCw size={36} className="animate-spin text-[#ffd025] mb-3" />
          <p className="text-sm font-bold">Cargando archivo de medios...</p>
        </div>
      ) : filteredImages.length === 0 ? (
        <div className="p-16 rounded-3xl bg-[#13131f]/60 border border-white/10 flex flex-col items-center justify-center text-center">
          <div className="w-20 h-20 rounded-3xl bg-white/5 border border-white/10 flex items-center justify-center mb-4 text-[#ffd025]">
            <FolderOpen size={36} />
          </div>
          <h3 className="text-lg font-black text-white uppercase tracking-wider mb-2">
            {searchQuery ? "No hay imágenes que coincidan" : "No has subido imágenes todavía"}
          </h3>
          <p className="text-xs text-gray-400 max-w-md mb-6">
            {searchQuery
              ? "Prueba buscando con otra palabra o borra el filtro de búsqueda."
              : "Selecciona varias imágenes de tus productos desde tu computadora o celular para tenerlas listas en tu galería."}
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
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-3.5 sm:gap-4">
          {filteredImages.map((img) => (
            <div
              key={img.id}
              className="group relative rounded-2xl bg-[#13131f] border border-white/10 overflow-hidden flex flex-col hover:border-[#ffd025] transition-all duration-200 hover:shadow-xl hover:shadow-[#ffd025]/10"
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
                      className="p-1.5 rounded-lg bg-black/80 hover:bg-[#ffd025] text-white hover:text-black transition"
                    >
                      {copiedId === img.id ? <Check size={14} /> : <Copy size={14} />}
                    </button>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleDeleteImage(img.id);
                      }}
                      title="Eliminar foto"
                      className="p-1.5 rounded-lg bg-red-600/80 hover:bg-red-600 text-white transition"
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>

                  <span className="text-[10px] text-[#ffd025] font-bold uppercase tracking-wider bg-black/70 px-2 py-0.5 rounded-full">
                    🔍 Ver grande
                  </span>
                </div>
              </div>

              {/* Info footer */}
              <div className="p-2.5 border-t border-white/5 space-y-1">
                <p className="text-[11px] font-bold text-white truncate" title={img.name}>
                  {img.name}
                </p>
                <div className="flex items-center justify-between text-[9px] text-gray-500">
                  <span>{new Date(img.createdAt).toLocaleDateString("es-CL")}</span>
                  {img.size ? <span>{Math.round(img.size / 1024)} KB</span> : null}
                </div>
              </div>
            </div>
          ))}
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
              <h4 className="text-sm sm:text-base font-bold text-white truncate max-w-md">
                {previewImage.name}
              </h4>
              <button
                onClick={() => setPreviewImage(null)}
                className="p-1.5 rounded-xl bg-white/10 hover:bg-white/20 text-white transition"
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
