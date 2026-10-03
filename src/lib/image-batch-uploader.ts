export async function compressImage(file: File, maxPx = 1200, quality = 0.82): Promise<Blob> {
  if (file.type === "image/svg+xml") return file;

  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(url);
      let { width, height } = img;
      if (width > maxPx || height > maxPx) {
        const ratio = Math.min(maxPx / width, maxPx / height);
        width = Math.round(width * ratio);
        height = Math.round(height * ratio);
      }
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        resolve(file);
        return;
      }
      ctx.drawImage(img, 0, 0, width, height);
      canvas.toBlob(
        (blob) => resolve(blob ?? file),
        "image/jpeg",
        quality
      );
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve(file);
    };
    img.src = url;
  });
}

export interface UploadBatchProgress {
  current: number;
  total: number;
  percent: number;
  fileName: string;
}

export async function uploadImagesBatch(
  files: File[],
  onProgress?: (progress: UploadBatchProgress) => void
): Promise<{ successCount: number; failedCount: number }> {
  let successCount = 0;
  let failedCount = 0;
  const total = files.length;
  const CONCURRENCY = 4;
  const apiBase = `${import.meta.env.BASE_URL}api`;

  let currentIndex = 0;

  async function worker() {
    while (currentIndex < total) {
      const index = currentIndex++;
      const file = files[index];
      try {
        if (onProgress) {
          onProgress({
            current: index + 1,
            total,
            percent: Math.round(((index + 1) / total) * 100),
            fileName: file.name,
          });
        }

        // 1. Compress file to ~80KB-150KB
        const blob = await compressImage(file);

        // 2. Request upload URL with original name
        const presignRes = await fetch(`${apiBase}/storage/uploads/request-url`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            name: file.name,
            size: blob.size,
            contentType: blob.type,
          }),
        });

        if (!presignRes.ok) throw new Error("Presign request failed");
        const { uploadURL } = await presignRes.json();

        // 3. Binary stream upload directly
        const putRes = await fetch(uploadURL, {
          method: "PUT",
          headers: { "Content-Type": blob.type || "image/jpeg" },
          body: blob,
        });

        if (!putRes.ok) throw new Error("PUT upload failed");
        successCount++;
      } catch (err) {
        console.error(`Error uploading image ${file.name}:`, err);
        failedCount++;
      }
    }
  }

  const workers = Array.from({ length: Math.min(CONCURRENCY, total) }, () => worker());
  await Promise.all(workers);

  return { successCount, failedCount };
}
