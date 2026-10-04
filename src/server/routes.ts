import express, { type Request, type Response } from "express";
import fs from "fs";
import path from "path";
import { randomUUID } from "crypto";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { GoogleGenAI } from "@google/genai";
import { dbManager, type Product, type Order, type OrderItem } from "./db.ts";

const JWT_SECRET = process.env.JWT_SECRET || "fellas-market-secret-jwt-key-2024";
const DATA_DIR = path.resolve(process.cwd(), "data");
const UPLOADS_DIR = path.join(DATA_DIR, "uploads");
const SEED_UPLOADS_DIR = path.resolve(process.cwd(), "src", "server", "seed_uploads");
const MEDIA_META_FILE = path.join(DATA_DIR, "media_library.json");
const SEED_MEDIA_FILE = path.resolve(process.cwd(), "src", "server", "seed_media.json");

if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}
if (!fs.existsSync(SEED_UPLOADS_DIR)) {
  fs.mkdirSync(SEED_UPLOADS_DIR, { recursive: true });
}

export const PRODUCT_IMAGE_SECTIONS = [
  "Promociones & Packs",
  "Piscos",
  "Cervezas",
  "Vinos & Espumantes",
  "Destilados & Licores",
  "Energizantes",
  "Bebidas & Gaseosas",
  "Jugos & Aguas",
  "Cigarros & Tabacos",
  "Snacks & Salados",
  "Dulces & Chocolates",
  "Hielo & Abarrotes",
  "Otros / General",
] as const;

export type ProductImageSection = (typeof PRODUCT_IMAGE_SECTIONS)[number];

export interface MediaMetaItem {
  id: string;
  name: string;
  url: string;
  createdAt: string;
  size?: number;
  category?: string;
  aiDetectedTitle?: string;
}

export function normalizeCategoryMatch(catStr: string, titleStr: string = ""): ProductImageSection {
  const catNorm = String(catStr || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const titleNorm = String(titleStr || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const combined = `${catNorm} ${titleNorm}`;

  if (!combined.trim()) return "Otros / General";

  // 1. Promociones & Packs (if combo/pack/promo)
  if (
    combined.includes("promo") ||
    combined.includes("pack") ||
    combined.includes("combo") ||
    combined.includes("oferta") ||
    (combined.includes("pisco") && combined.includes("bebida")) ||
    (combined.includes("whisky") && combined.includes("bebida")) ||
    (combined.includes("ron") && combined.includes("bebida")) ||
    (combined.includes("vodka") && combined.includes("bebida"))
  ) {
    return "Promociones & Packs";
  }

  // 2. Cervezas (MUST check before generic drinks or waters)
  if (
    combined.includes("cerveza") ||
    combined.includes("beer") ||
    combined.includes("schop") ||
    combined.includes("sixpack") ||
    combined.includes("twelvepack") ||
    combined.includes("cristal") ||
    combined.includes("escudo") ||
    combined.includes("heineken") ||
    combined.includes("stella") ||
    combined.includes("budweiser") ||
    combined.includes("coronita") ||
    combined.includes("corona") ||
    combined.includes("becker") ||
    combined.includes("kunstmann") ||
    combined.includes("austral") ||
    combined.includes("royal guard") ||
    combined.includes("kross") ||
    combined.includes("baltica") ||
    combined.includes("lager") ||
    combined.includes("ipa") ||
    combined.includes("pilsen") ||
    combined.includes("cuello negro")
  ) {
    return "Cervezas";
  }

  // 3. Piscos (MUST check before generic drinks or waters)
  if (
    combined.includes("pisco") ||
    combined.includes("mistral") ||
    combined.includes("alto del carmen") ||
    combined.includes("altodelcarmen") ||
    combined.includes("campanario") ||
    combined.includes("capel") ||
    combined.includes("horcon") ||
    combined.includes("bauza") ||
    combined.includes("control c") ||
    combined.includes("malpaso") ||
    combined.includes("tres erres")
  ) {
    return "Piscos";
  }

  // 4. Vinos & Espumantes
  if (
    combined.includes("vino") ||
    combined.includes("wine") ||
    combined.includes("espumante") ||
    combined.includes("champagne") ||
    combined.includes("prosecco") ||
    combined.includes("cava") ||
    combined.includes("tinto") ||
    combined.includes("blanco") ||
    combined.includes("carmenere") ||
    combined.includes("cabernet") ||
    combined.includes("merlot") ||
    combined.includes("chardonnay") ||
    combined.includes("sauvignon") ||
    combined.includes("casillero") ||
    combined.includes("gato negro") ||
    combined.includes("gatonegro") ||
    combined.includes("santa rita") ||
    combined.includes("concha y toro") ||
    combined.includes("undurraga") ||
    combined.includes("valdivieso") ||
    combined.includes("riccadonna") ||
    combined.includes("chandon") ||
    combined.includes("misiones de rengo")
  ) {
    return "Vinos & Espumantes";
  }

  // 5. Destilados & Licores
  if (
    combined.includes("whisky") ||
    combined.includes("whiskey") ||
    combined.includes("ron") ||
    combined.includes("vodka") ||
    combined.includes("gin") ||
    combined.includes("tequila") ||
    combined.includes("licor") ||
    combined.includes("destilado") ||
    combined.includes("fernet") ||
    combined.includes("jager") ||
    combined.includes("baileys") ||
    combined.includes("aperol") ||
    combined.includes("campari") ||
    combined.includes("ramazzotti") ||
    combined.includes("absolut") ||
    combined.includes("smirnoff") ||
    combined.includes("jack daniel") ||
    combined.includes("johnnie") ||
    combined.includes("chivas") ||
    combined.includes("ballantine") ||
    combined.includes("bacardi") ||
    combined.includes("havana")
  ) {
    return "Destilados & Licores";
  }

  // 6. Energizantes
  if (
    combined.includes("energiz") ||
    combined.includes("energetica") ||
    combined.includes("energy") ||
    combined.includes("monster") ||
    combined.includes("red bull") ||
    combined.includes("redbull") ||
    combined.includes("score") ||
    combined.includes("mr big") ||
    combined.includes("mrbig") ||
    combined.includes("volt") ||
    combined.includes("rockstar")
  ) {
    return "Energizantes";
  }

  // 7. Cigarros & Tabacos
  if (
    combined.includes("cigarro") ||
    combined.includes("cigar") ||
    combined.includes("tabaco") ||
    combined.includes("vape") ||
    combined.includes("pod") ||
    combined.includes("pall mall") ||
    combined.includes("lucky") ||
    combined.includes("marlboro") ||
    combined.includes("kent") ||
    combined.includes("dunhill") ||
    combined.includes("belmont") ||
    combined.includes("papelillo") ||
    combined.includes("encendedor")
  ) {
    return "Cigarros & Tabacos";
  }

  // 8. Snacks & Salados
  if (
    combined.includes("snack") ||
    combined.includes("papa") ||
    combined.includes("lays") ||
    combined.includes("ramita") ||
    combined.includes("dorito") ||
    combined.includes("cheeto") ||
    combined.includes("mani") ||
    combined.includes("frutos secos") ||
    combined.includes("aceituna") ||
    combined.includes("piqueo") ||
    combined.includes("kryzpo") ||
    combined.includes("pringles") ||
    combined.includes("tika")
  ) {
    return "Snacks & Salados";
  }

  // 9. Dulces & Chocolates
  if (
    combined.includes("dulce") ||
    combined.includes("chocolate") ||
    combined.includes("galleta") ||
    combined.includes("gomita") ||
    combined.includes("caramelo") ||
    combined.includes("alfajor") ||
    combined.includes("sahne") ||
    combined.includes("vizzio") ||
    combined.includes("trencito") ||
    combined.includes("super 8") ||
    combined.includes("chocman") ||
    combined.includes("oreo") ||
    combined.includes("triton") ||
    combined.includes("golosina")
  ) {
    return "Dulces & Chocolates";
  }

  // 10. Jugos & Aguas (ONLY if truly juice or water)
  if (
    combined.includes("jugo") ||
    combined.includes("nectar") ||
    combined.includes("watts") ||
    combined.includes("andina") ||
    combined.includes("vivo") ||
    combined.includes("agua mineral") ||
    combined.includes("cachantun") ||
    combined.includes("benedicto") ||
    combined.includes("vital") ||
    combined.includes("puyehue") ||
    combined.includes("agua sin gas") ||
    combined.includes("agua con gas")
  ) {
    return "Jugos & Aguas";
  }

  // 11. Bebidas & Gaseosas
  if (
    combined.includes("bebida") ||
    combined.includes("gaseosa") ||
    combined.includes("soda") ||
    combined.includes("coca") ||
    combined.includes("sprite") ||
    combined.includes("fanta") ||
    combined.includes("pepsi") ||
    combined.includes("kem") ||
    combined.includes("bilz") ||
    combined.includes("pap") ||
    combined.includes("ginger") ||
    combined.includes("tonica")
  ) {
    return "Bebidas & Gaseosas";
  }

  // 12. Hielo & Abarrotes
  if (
    combined.includes("hielo") ||
    combined.includes("ice") ||
    combined.includes("carbon") ||
    combined.includes("vaso") ||
    combined.includes("servilleta") ||
    combined.includes("abarrote") ||
    combined.includes("despensa")
  ) {
    return "Hielo & Abarrotes";
  }

  return "Otros / General";
}

export function classifyImageHeuristic(filename: string): { category: ProductImageSection; title?: string } {
  const norm = String(filename || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  // Promociones & Packs (ej: combo pisco + bebida + hielo, pack promo)
  if (
    norm.includes("pack") ||
    norm.includes("combo") ||
    norm.includes("promo") ||
    norm.includes("oferta") ||
    (norm.includes("pisco") && norm.includes("bebida")) ||
    (norm.includes("whisky") && norm.includes("bebida")) ||
    (norm.includes("ron") && norm.includes("bebida")) ||
    (norm.includes("vodka") && norm.includes("bebida")) ||
    (norm.includes("hielo") && (norm.includes("pisco") || norm.includes("bebida") || norm.includes("whisky")))
  ) {
    return { category: "Promociones & Packs" };
  }

  // Piscos
  if (
    norm.includes("pisco") ||
    norm.includes("mistral") ||
    norm.includes("alto del carmen") ||
    norm.includes("altodelcarmen") ||
    norm.includes("campanario") ||
    norm.includes("capel") ||
    norm.includes("horcon") ||
    norm.includes("tres erres") ||
    norm.includes("3 erres") ||
    norm.includes("control c") ||
    norm.includes("bauza") ||
    norm.includes("mal paso") ||
    norm.includes("waqar")
  ) {
    return { category: "Piscos" };
  }

  // Energizantes
  if (
    norm.includes("monster") ||
    norm.includes("red bull") ||
    norm.includes("redbull") ||
    norm.includes("score") ||
    norm.includes("mr big") ||
    norm.includes("mrbig") ||
    norm.includes("volt") ||
    norm.includes("rockstar") ||
    norm.includes("energiz") ||
    norm.includes("energy")
  ) {
    return { category: "Energizantes" };
  }

  // Jugos & Aguas
  if (
    norm.includes("jugo") ||
    norm.includes("nectar") ||
    norm.includes("watts") ||
    norm.includes("andina") ||
    norm.includes("vivo") ||
    norm.includes("agua") ||
    norm.includes("water") ||
    norm.includes("cachantun") ||
    norm.includes("benedicto") ||
    norm.includes("vital") ||
    norm.includes("puyehue") ||
    norm.includes("mas agua") ||
    norm.includes("aquarius")
  ) {
    return { category: "Jugos & Aguas" };
  }

  // Bebidas & Gaseosas
  if (
    norm.includes("bebida") ||
    norm.includes("gaseosa") ||
    norm.includes("soda") ||
    norm.includes("coca") ||
    norm.includes("sprite") ||
    norm.includes("fanta") ||
    norm.includes("pepsi") ||
    norm.includes("7up") ||
    norm.includes("kem") ||
    norm.includes("pap") ||
    norm.includes("bilz") ||
    norm.includes("nordic") ||
    norm.includes("ginger") ||
    norm.includes("tonica") ||
    norm.includes("crush") ||
    norm.includes("quatro") ||
    norm.includes("limon soda") ||
    norm.includes("limonsoda") ||
    norm.includes("canada dry")
  ) {
    return { category: "Bebidas & Gaseosas" };
  }

  // Cervezas
  if (
    norm.includes("cerveza") ||
    norm.includes("beer") ||
    norm.includes("cristal") ||
    norm.includes("escudo") ||
    norm.includes("corona") ||
    norm.includes("coronita") ||
    norm.includes("heineken") ||
    norm.includes("stella") ||
    norm.includes("budweiser") ||
    norm.includes("royal") ||
    norm.includes("austral") ||
    norm.includes("kunstmann") ||
    norm.includes("becker") ||
    norm.includes("cusquena") ||
    norm.includes("sol") ||
    norm.includes("miller") ||
    norm.includes("coors") ||
    norm.includes("kross") ||
    norm.includes("baltica") ||
    norm.includes("schop") ||
    norm.includes("ipa") ||
    norm.includes("lager") ||
    norm.includes("golden") ||
    norm.includes("torobayo") ||
    norm.includes("cuello negro")
  ) {
    return { category: "Cervezas" };
  }

  // Vinos & Espumantes
  if (
    norm.includes("vino") ||
    norm.includes("wine") ||
    norm.includes("espumante") ||
    norm.includes("champagne") ||
    norm.includes("prosecco") ||
    norm.includes("cava") ||
    norm.includes("tinto") ||
    norm.includes("blanco") ||
    norm.includes("cabernet") ||
    norm.includes("carmenere") ||
    norm.includes("merlot") ||
    norm.includes("chardonnay") ||
    norm.includes("sauvignon") ||
    norm.includes("casillero") ||
    norm.includes("gato negro") ||
    norm.includes("gatonegro") ||
    norm.includes("santa rita") ||
    norm.includes("concha y toro") ||
    norm.includes("san pedro") ||
    norm.includes("tarapaca") ||
    norm.includes("undurraga") ||
    norm.includes("valdivieso") ||
    norm.includes("riccadonna") ||
    norm.includes("chandon") ||
    norm.includes("misiones de rengo") ||
    norm.includes("castillo de molina") ||
    norm.includes("clos")
  ) {
    return { category: "Vinos & Espumantes" };
  }

  // Destilados & Licores
  if (
    norm.includes("whisky") ||
    norm.includes("whiskey") ||
    norm.includes("ron") ||
    norm.includes("vodka") ||
    norm.includes("gin") ||
    norm.includes("tequila") ||
    norm.includes("licor") ||
    norm.includes("fernet") ||
    norm.includes("jagermeister") ||
    norm.includes("jager") ||
    norm.includes("baileys") ||
    norm.includes("aperol") ||
    norm.includes("campari") ||
    norm.includes("ramazzotti") ||
    norm.includes("absolut") ||
    norm.includes("smirnoff") ||
    norm.includes("grey goose") ||
    norm.includes("havana") ||
    norm.includes("bacardi") ||
    norm.includes("barcelo") ||
    norm.includes("jack daniel") ||
    norm.includes("johnnie") ||
    norm.includes("chivas") ||
    norm.includes("ballantine") ||
    norm.includes("red label") ||
    norm.includes("black label") ||
    norm.includes("tanqueray") ||
    norm.includes("beefeater") ||
    norm.includes("bombay") ||
    norm.includes("jose cuervo") ||
    norm.includes("malibu") ||
    norm.includes("st germain")
  ) {
    return { category: "Destilados & Licores" };
  }

  // Cigarros & Tabacos
  if (
    norm.includes("cigarro") ||
    norm.includes("cigarros") ||
    norm.includes("cigarrette") ||
    norm.includes("cigar") ||
    norm.includes("tabaco") ||
    norm.includes("vape") ||
    norm.includes("pod") ||
    norm.includes("pall mall") ||
    norm.includes("pallmall") ||
    norm.includes("lucky") ||
    norm.includes("luckies") ||
    norm.includes("kent") ||
    norm.includes("marlboro") ||
    norm.includes("dunhill") ||
    norm.includes("belmont") ||
    norm.includes("rothmans") ||
    norm.includes("sedas") ||
    norm.includes("ocb") ||
    norm.includes("papelillo") ||
    norm.includes("encendedor") ||
    norm.includes("clipper") ||
    norm.includes("bic")
  ) {
    return { category: "Cigarros & Tabacos" };
  }

  // Snacks & Salados
  if (
    norm.includes("papa") ||
    norm.includes("papas") ||
    norm.includes("lays") ||
    norm.includes("ramitas") ||
    norm.includes("doritos") ||
    norm.includes("cheetos") ||
    norm.includes("chizito") ||
    norm.includes("mani") ||
    norm.includes("frutos secos") ||
    norm.includes("pistacho") ||
    norm.includes("almendra") ||
    norm.includes("aceituna") ||
    norm.includes("nacho") ||
    norm.includes("tika") ||
    norm.includes("kryzpo") ||
    norm.includes("pringles") ||
    norm.includes("snack") ||
    norm.includes("gatolate")
  ) {
    return { category: "Snacks & Salados" };
  }

  // Dulces & Chocolates
  if (
    norm.includes("chocolate") ||
    norm.includes("dulce") ||
    norm.includes("galleta") ||
    norm.includes("gomita") ||
    norm.includes("caramelo") ||
    norm.includes("sahne nuss") ||
    norm.includes("sahnenuss") ||
    norm.includes("vizzio") ||
    norm.includes("trencito") ||
    norm.includes("costa rama") ||
    norm.includes("super 8") ||
    norm.includes("super8") ||
    norm.includes("chocman") ||
    norm.includes("m&m") ||
    norm.includes("snickers") ||
    norm.includes("kit kat") ||
    norm.includes("mentita") ||
    norm.includes("halls") ||
    norm.includes("chicle") ||
    norm.includes("alfajor") ||
    norm.includes("fruggele") ||
    norm.includes("triton") ||
    norm.includes("oreo") ||
    norm.includes("kuky")
  ) {
    return { category: "Dulces & Chocolates" };
  }

  // Hielo & Abarrotes
  if (
    norm.includes("hielo") ||
    norm.includes("ice") ||
    norm.includes("carbon") ||
    norm.includes("vaso") ||
    norm.includes("servilleta") ||
    norm.includes("abarrote") ||
    norm.includes("despensa") ||
    norm.includes("arroz") ||
    norm.includes("aceite")
  ) {
    return { category: "Hielo & Abarrotes" };
  }

  return { category: "Otros / General" };
}

export async function classifyImageWithAI(
  buffer: Buffer,
  filename: string,
  mimeType = "image/jpeg"
): Promise<{ category: ProductImageSection; title?: string }> {
  const heuristic = classifyImageHeuristic(filename);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey || buffer.length === 0) {
    return heuristic;
  }

  try {
    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
    const base64Data = buffer.toString("base64");

    const prompt = `Eres un sistema experto en visión artificial para botillerías, minimarkets y tiendas de conveniencia en Chile y Latinoamérica.
Tu tarea es inspeccionar exhaustivamente la imagen proporcionada (etiqueta, marca, forma del envase, lata, botella, paquete o combo) y clasificar el producto en EXACTAMENTE UNA de estas categorías oficiales:

CATEGORÍAS OFICIALES PERMITIDAS:
1. "Promociones & Packs" -> Combos y ofertas (ej: Pisco + Bebida + Hielo, Whisky + Bebida + Hielo, Vodka + Energética + Hielo, packs de cerveza + vaso, etc.).
2. "Piscos" -> Botellas o latas individuales de pisco (Mistral, Alto del Carmen, Campanario, Capel, Horcón Quemado, Bauzá, Control C, MalPaso, Waqar, etc.).
3. "Cervezas" -> Cervezas individuales, six-packs o twelve-packs (Cristal, Escudo, Corona, Coronita, Heineken, Stella Artois, Royal Guard, Becker, Kunstmann, Austral, Budweiser, Kross, Baltica, etc.).
4. "Vinos & Espumantes" -> Vinos tintos, blancos, rosé, botellones, cajas de vino, espumantes, champagnes, prosecco (Casillero del Diablo, Gato Negro, Concha y Toro, Santa Helena, Undurraga, Valdivieso, Chandon, Riccadonna, etc.).
5. "Destilados & Licores" -> Whisky, Ron, Vodka, Gin, Tequila, Fernet, Jägermeister, Baileys, Aperol, Campari, Ramazzotti, Absolut, Jack Daniel's, Johnnie Walker, etc.
6. "Energizantes" -> Bebidas energéticas individuales o packs (Monster Energy, Red Bull, Score, Mr Big, Rockstar, Volt, etc.).
7. "Bebidas & Gaseosas" -> Gaseosas y refrescos carbonatados (Coca-Cola, Sprite, Fanta, Pepsi, Kem Piña, Bilz, Pap, Nordic Ginger Ale, Agua Tónica, Canada Dry, etc.).
8. "Jugos & Aguas" -> Jugos en caja/botella (Watts, Andina, Vivo), aguas minerales con y sin gas (Cachantun, Benedictino, Vital, Puyehue), aguas saborizadas, etc.
9. "Cigarros & Tabacos" -> Cajetillas de cigarrillos, tabacos para enrolar, vapes, pods electrónicos, filtros, sedas/papelillos (Lucky Strike, Pall Mall, Kent, Marlboro, Dunhill, OCB, etc.).
10. "Snacks & Salados" -> Papas fritas (Lay's, Tika, Kryzpo, Pringles), Ramitas (queso, sal), Doritos, Cheetos, Gatolate, maní salado/japonés, frutos secos, aceitunas, nachos.
11. "Dulces & Chocolates" -> Chocolates (Sahne Nuss, Vizzio, Costa Rama, Trencito), galletas (Oreo, Tritón, Kuky), gomitas (Fruggele), caramelos, alfajores, Super 8, Chocman, bombones, chicles.
12. "Hielo & Abarrotes" -> Bolsas de hielo en cubos o frappé, carbón, vasos plásticos, servilletas, abarrotes básicos de despensa.
13. "Otros / General" -> Cualquier otro artículo que no encaje con precisión en las categorías anteriores.

Nombre original del archivo sugerido: "${filename}"

REGLAS OBLIGATORIAS:
- Si en la imagen se observa un combo de botella con bebida o hielo, clasifica siempre como "Promociones & Packs".
- Si es una bebida energética (Monster, Red Bull, Score, etc.), clasifícala estrictamente como "Energizantes".
- Si es un jugo o agua mineral, clasifícala estrictamente como "Jugos & Aguas".
- Si son cigarros, vapes o tabaco, clasifícalo estrictamente como "Cigarros & Tabacos".
- Si son papas fritas o snacks salados, clasifícalos como "Snacks & Salados".
- Si son chocolates, galletas o dulces, clasifícalos como "Dulces & Chocolates".

Responde ÚNICAMENTE en formato JSON válido:
{
  "category": "Nombre exacto de la categoría oficial",
  "title": "Nombre comercial descriptivo y preciso del producto identificado (ej: Pisco Mistral 35° 750cc, Pack Pisco Alto del Carmen + Coca Cola + Hielo, Energética Monster Mango Loco 473ml, Papas Lay's Corte Liso 200g, etc.)"
}`;

    const response = await ai.models.generateContent({
      model: "gemini-3.8-flash",
      contents: [
        {
          role: "user",
          parts: [
            {
              inlineData: {
                data: base64Data,
                mimeType: mimeType.includes("png") ? "image/png" : mimeType.includes("webp") ? "image/webp" : "image/jpeg",
              },
            },
            { text: prompt },
          ],
        },
      ],
      config: { responseMimeType: "application/json" },
    });

    const parsed = JSON.parse(response.text || "{}");
    const matchedCategory = normalizeCategoryMatch(parsed.category || "", parsed.title || "");
    if (matchedCategory) {
      return {
        category: matchedCategory,
        title: parsed.title || heuristic.title,
      };
    }
  } catch (err) {
    console.warn("Gemini AI image classification fallback:", err);
  }

  return heuristic;
}

export function getMediaLibraryMeta(): MediaMetaItem[] {
  try {
    if (fs.existsSync(MEDIA_META_FILE)) {
      const items = JSON.parse(fs.readFileSync(MEDIA_META_FILE, "utf-8"));
      if (Array.isArray(items) && items.length > 0) return items;
    }
  } catch {}

  try {
    if (fs.existsSync(SEED_MEDIA_FILE)) {
      const items = JSON.parse(fs.readFileSync(SEED_MEDIA_FILE, "utf-8"));
      if (Array.isArray(items) && items.length > 0) {
        saveMediaLibraryMeta(items);
        return items;
      }
    }
  } catch {}

  return [];
}

export function saveMediaLibraryMeta(items: MediaMetaItem[]) {
  try {
    fs.writeFileSync(MEDIA_META_FILE, JSON.stringify(items, null, 2), "utf-8");
    try {
      fs.writeFileSync(SEED_MEDIA_FILE, JSON.stringify(items, null, 2), "utf-8");
    } catch {}
  } catch (err) {
    console.error("Error saving media library meta:", err);
  }
}

export function saveExtractedImage(
  imgBuffer: Buffer,
  originalName: string,
  fallbackExt = "jpg"
): { fileId: string; url: string; category: string } {
  let ext = fallbackExt;
  const extMatch = originalName.match(/\.(png|jpe?g|webp|gif|bmp)$/i);
  if (extMatch) {
    ext = extMatch[1].toLowerCase().replace("jpeg", "jpg");
  } else if (imgBuffer.length > 4) {
    if (imgBuffer[0] === 0x89 && imgBuffer[1] === 0x50 && imgBuffer[2] === 0x4e && imgBuffer[3] === 0x47) {
      ext = "png";
    } else if (imgBuffer[0] === 0xff && imgBuffer[1] === 0xd8 && imgBuffer[2] === 0xff) {
      ext = "jpg";
    } else if (imgBuffer[0] === 0x52 && imgBuffer[1] === 0x49 && imgBuffer[2] === 0x46 && imgBuffer[3] === 0x46) {
      ext = "webp";
    } else if (imgBuffer[0] === 0x47 && imgBuffer[1] === 0x49 && imgBuffer[2] === 0x46) {
      ext = "gif";
    }
  }

  const fileId = `${randomUUID()}.${ext}`;
  const filePath = path.join(UPLOADS_DIR, fileId);
  const seedPath = path.join(SEED_UPLOADS_DIR, fileId);

  fs.writeFileSync(filePath, imgBuffer);
  try {
    fs.writeFileSync(seedPath, imgBuffer);
  } catch {}

  const detected = classifyImageHeuristic(originalName);

  try {
    const metaList = getMediaLibraryMeta();
    metaList.unshift({
      id: fileId,
      name: originalName || `Excel_${fileId.slice(0, 8)}.${ext}`,
      url: `/api/storage/objects/${fileId}`,
      createdAt: new Date().toISOString(),
      size: imgBuffer.length,
      category: detected.category,
      aiDetectedTitle: detected.title,
    });
    saveMediaLibraryMeta(metaList);
  } catch (err) {
    console.error("Error updating media library meta for excel image:", err);
  }

  return {
    fileId,
    url: `/api/storage/objects/${fileId}`,
    category: detected.category,
  };
}

export async function extractImagesFromXlsx(buffer: Buffer) {
  const rowImages: { [excelRow: number]: { filename: string; buffer: Buffer } } = {};
  const cellImages: { [key: string]: { filename: string; buffer: Buffer } } = {};
  const allMediaList: { filename: string; buffer: Buffer }[] = [];
  const allMediaByName: { [name: string]: { filename: string; buffer: Buffer } } = {};

  try {
    const zip = await JSZip.loadAsync(buffer);

    for (const [filepath, file] of Object.entries(zip.files)) {
      if (filepath.startsWith("xl/media/") && !file.dir) {
        const filename = filepath.replace("xl/media/", "");
        const content = await file.async("nodebuffer");
        const item = { filename, buffer: content };
        allMediaList.push(item);
        allMediaByName[filename.toLowerCase()] = item;
      }
    }

    allMediaList.sort((a, b) => {
      const numA = parseInt(a.filename.replace(/\D/g, ""), 10) || 0;
      const numB = parseInt(b.filename.replace(/\D/g, ""), 10) || 0;
      return numA - numB;
    });

    const relsMap: { [relsFile: string]: { [rId: string]: string } } = {};
    for (const [filepath, file] of Object.entries(zip.files)) {
      if (filepath.endsWith(".rels") && !file.dir) {
        const xml = await file.async("text");
        const map: { [rId: string]: string } = {};
        const relMatches = xml.matchAll(/<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g);
        for (const m of relMatches) {
          const rId = m[1];
          let target = m[2];
          const parts = target.split("/");
          map[rId] = parts[parts.length - 1];
        }
        relsMap[filepath] = map;
      }
    }

    for (const [filepath, file] of Object.entries(zip.files)) {
      if (filepath.startsWith("xl/drawings/drawing") && filepath.endsWith(".xml")) {
        const xml = await file.async("text");
        const relsPath = filepath.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
        const rels = relsMap[relsPath] || {};

        const anchorRegex = /<xdr:(?:twoCellAnchor|oneCellAnchor)[^>]*>([\s\S]*?)<\/xdr:(?:twoCellAnchor|oneCellAnchor)>/g;
        let match;
        while ((match = anchorRegex.exec(xml)) !== null) {
          const block = match[1];
          const fromRowMatch = block.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
          const fromColMatch = block.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/);
          const blipMatch = block.match(/<a:blip[^>]+(?:r:embed|r:link)="([^"]+)"/);

          if (fromRowMatch && blipMatch) {
            const row = parseInt(fromRowMatch[1], 10);
            const rId = blipMatch[1];
            const mediaName = rels[rId];
            if (mediaName && allMediaByName[mediaName.toLowerCase()]) {
              const mediaObj = allMediaByName[mediaName.toLowerCase()];
              rowImages[row] = mediaObj;
              if (fromColMatch) {
                const col = parseInt(fromColMatch[1], 10);
                cellImages[`${row},${col}`] = mediaObj;
              }
            }
          }
        }
      }
    }

    for (const [filepath, file] of Object.entries(zip.files)) {
      if (filepath.includes("cellimages.xml") && !filepath.endsWith(".rels") && !file.dir) {
        const xml = await file.async("text");
        const relsPath = filepath.replace("cellimages.xml", "_rels/cellimages.xml.rels");
        const rels = relsMap[relsPath] || relsMap["xl/_rels/cellimages.xml.rels"] || {};

        const blipMatches = xml.matchAll(/<a:blip[^>]+(?:r:embed|r:link)="([^"]+)"/g);
        let idx = 0;
        for (const bm of blipMatches) {
          const rId = bm[1];
          const mediaName = rels[rId];
          if (mediaName && allMediaByName[mediaName.toLowerCase()]) {
            const mediaObj = allMediaByName[mediaName.toLowerCase()];
            if (!rowImages[idx + 1]) {
              rowImages[idx + 1] = mediaObj;
            }
          }
          idx++;
        }
      }
    }
  } catch (err) {
    console.error("Error reading xlsx images with JSZip:", err);
  }

  return { rowImages, cellImages, allMediaList, allMediaByName };
}

export const apiRouter = express.Router();

// Health check & Keep-alive endpoint
apiRouter.get("/health", (_req, res) => {
  res.json({
    status: "ok",
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    keepAlive: true,
  });
});

apiRouter.get("/ping", (_req, res) => {
  res.json({ pong: true, timestamp: Date.now() });
});

// In-memory cache for ip -> geo
const ipGeoCache = new Map<string, { city: string; region: string; country: string }>();

async function lookupIpGeo(ip: string) {
  if (
    !ip ||
    ip === "::1" ||
    ip.startsWith("127.") ||
    ip.startsWith("10.") ||
    ip.startsWith("192.168.") ||
    ip.startsWith("172.")
  ) {
    return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
  }
  const cached = ipGeoCache.get(ip);
  if (cached) return cached;
  try {
    const res = await fetch(
      `http://ip-api.com/json/${encodeURIComponent(ip)}?fields=status,country,regionName,city`,
      { signal: AbortSignal.timeout(2500) }
    );
    if (!res.ok) return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
    const data = (await res.json()) as { status?: string; country?: string; regionName?: string; city?: string };
    if (data.status !== "success") return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
    const geo = {
      city: data.city || "Puerto Montt",
      region: data.regionName || "Los Lagos",
      country: data.country || "Chile",
    };
    ipGeoCache.set(ip, geo);
    return geo;
  } catch {
    return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
  }
}

function getCustomerIdFromReq(req: Request): number | null {
  const auth = req.headers.authorization;
  if (!auth?.startsWith("Bearer ")) return null;
  try {
    const token = auth.slice(7);
    const decoded = jwt.verify(token, JWT_SECRET) as { customerId: number };
    return decoded.customerId ?? null;
  } catch {
    return null;
  }
}

// 1. Health
apiRouter.get("/healthz", (_req, res) => {
  res.json({ status: "ok" });
});

// 2. Menu
apiRouter.get("/menu", (_req, res) => {
  const settings = dbManager.getSettings();
  const categories = dbManager.getCategories();
  const aisles = dbManager.getAisles();
  const subcategories = dbManager.getSubcategories();
  const allProducts = dbManager.getProducts();
  const orders = dbManager.getOrders();

  // Compute bestsellers from last 30 days orders
  const thirtyDaysAgo = Date.now() - 30 * 24 * 3600 * 1000;
  const productQtyMap = new Map<string, number>();

  for (const order of orders) {
    if (order.status === "Anulado") continue;
    if (new Date(order.createdAt).getTime() < thirtyDaysAgo) continue;
    for (const it of order.items || []) {
      const q = Number(it.quantity) || 1;
      productQtyMap.set(it.name, (productQtyMap.get(it.name) || 0) + q);
    }
  }

  const topNames = new Set(
    [...productQtyMap.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name]) => name)
  );

  const products = allProducts
    .filter((p) => !p.hidden)
    .map((p) => ({
      ...p,
      bestseller: topNames.has(p.name) || p.oferta,
    }));

  res.json({
    settings,
    categories,
    aisles,
    subcategories,
    products,
  });
});

// 3. Settings
apiRouter.get("/settings", (_req, res) => {
  res.json(dbManager.getSettings());
});

apiRouter.patch("/settings", (req, res) => {
  const updated = dbManager.updateSettings(req.body);
  res.json(updated);
});

// Google Maps Reviews API
apiRouter.get("/google-maps-reviews", (_req, res) => {
  const settings = dbManager.getSettings();
  res.json({
    rating: settings.googleMapsRating || 4.9,
    reviewsCount: settings.googleMapsReviewsCount || 142,
    reviews: settings.googleMapsReviews || []
  });
});

// 4. Categories
apiRouter.get("/categories", (_req, res) => {
  res.json(dbManager.getCategories());
});

apiRouter.post("/categories", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addCategory(name);
  res.status(201).json(created);
});

apiRouter.delete("/categories/:id", (req, res) => {
  const ok = dbManager.deleteCategory(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/categories/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderCategories(ids);
  res.json({ ok: true });
});

// 5. Aisles
apiRouter.get("/aisles", (_req, res) => {
  res.json(dbManager.getAisles());
});

apiRouter.post("/aisles", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addAisle(name);
  res.status(201).json(created);
});

apiRouter.patch("/aisles/:id", (req, res) => {
  const updated = dbManager.updateAisle(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Aisle not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/aisles/:id", (req, res) => {
  const ok = dbManager.deleteAisle(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/aisles/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderAisles(ids);
  res.json({ ok: true });
});

// 6. Subcategories
apiRouter.get("/subcategories", (_req, res) => {
  res.json(dbManager.getSubcategories());
});

apiRouter.post("/subcategories", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addSubcategory(name);
  res.status(201).json(created);
});

apiRouter.delete("/subcategories/:id", (req, res) => {
  const ok = dbManager.deleteSubcategory(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/subcategories/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderSubcategories(ids);
  res.json({ ok: true });
});

// 7. Products
apiRouter.get("/products", (_req, res) => {
  res.json(dbManager.getProducts());
});

apiRouter.post("/products", (req, res) => {
  const body = req.body;
  const created = dbManager.addProduct({
    name: body.name,
    price: Number(body.price) || 0,
    image: body.image || "",
    category: body.category || "",
    aisle: body.aisle || "",
    subcategory: body.subcategory || "",
    optionsTitle: body.optionsTitle || "",
    options: Array.isArray(body.options) ? body.options : [],
    oferta: Boolean(body.oferta),
    depositoEnabled: Boolean(body.depositoEnabled),
    depositoAmount: Number(body.depositoAmount) || 0,
    position: dbManager.getProducts().length,
    publishedSocial: Boolean(body.publishedSocial),
    hidden: Boolean(body.hidden),
    transferenciaEnabled: Boolean(body.transferenciaEnabled),
    transferenciaAmount: Number(body.transferenciaAmount) || 0,
    contingencyEnabled: Boolean(body.contingencyEnabled),
  });
  res.status(201).json(created);
});

apiRouter.post("/admin/products/contingency-batch", (req, res) => {
  const { ids = [], enable = true } = req.body;
  const idsSet = new Set((ids as number[]).map(Number));
  const products = dbManager.getProducts();
  for (const p of products) {
    if (idsSet.has(p.id)) {
      dbManager.updateProduct(p.id, { contingencyEnabled: Boolean(enable) });
    }
  }
  res.json({ ok: true });
});

apiRouter.patch("/products/:id", (req, res) => {
  const updated = dbManager.updateProduct(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Product not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/products/:id", (req, res) => {
  const ok = dbManager.deleteProduct(Number(req.params.id));
  res.json({ ok });
});

apiRouter.delete("/admin/products/all", (_req, res) => {
  dbManager.deleteAllProducts();
  res.json({ ok: true });
});

apiRouter.get("/admin/download-template", (_req, res) => {
  try {
    const data = [
      { Nombre: "Cerveza Cristal 350cc", Pasillo: "Cervezas", Categoria: "Latas", Subcategoria: "Nacionales", Precio: 1200, Oferta: "No", Imagen: "https://images.unsplash.com/photo-1608270127590-27f991f89cd0?w=600&auto=format&fit=crop&q=80" },
      { Nombre: "Pisco Mistral 35° 750cc", Pasillo: "Licores y Destilados", Categoria: "Piscos", Subcategoria: "Piscos 35°", Precio: 7990, Oferta: "Si", Imagen: "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80" },
      { Nombre: "Bebida Coca Cola 1.5L", Pasillo: "Bebidas y Jugos", Categoria: "Bebidas Gaseosas", Subcategoria: "Descartables", Precio: 2100, Oferta: "No", Imagen: "https://images.unsplash.com/photo-1622483767028-3f66f32aef97?w=600&auto=format&fit=crop&q=80" }
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Productos");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Disposition", 'attachment; filename="plantilla_productos_fellas.xlsx"');
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buf);
  } catch (err: any) {
    res.status(500).json({ error: "Error generating template" });
  }
});

// Exportar catálogo completo a Excel con imágenes incrustadas en máxima calidad
apiRouter.get("/admin/export-excel-with-images", async (_req, res) => {
  try {
    const products = dbManager.getProducts();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Fellas Market";
    workbook.created = new Date();
    
    const worksheet = workbook.addWorksheet("Productos", {
      views: [{ showGridLines: true }]
    });

    worksheet.columns = [
      { header: "Nombre", key: "name", width: 38 },
      { header: "Pasillo", key: "aisle", width: 25 },
      { header: "Categoria", key: "category", width: 25 },
      { header: "Subcategoria", key: "subcategory", width: 25 },
      { header: "Precio", key: "price", width: 16 },
      { header: "Oferta", key: "oferta", width: 12 },
      { header: "Imagen", key: "image", width: 20 },
    ];

    // Formato de cabeceras
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF181828" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 30;

    for (let i = 0; i < products.length; i++) {
      const p = products[i];
      const excelRowIdx = i + 2; // 1-indexed

      const row = worksheet.addRow({
        name: p.name || "",
        aisle: p.aisle || "General",
        category: p.category || "General",
        subcategory: p.subcategory || "General",
        price: p.price || 0,
        oferta: p.oferta ? "Si" : "No",
        image: "",
      });

      row.height = 70;
      row.alignment = { vertical: "middle", horizontal: "left" };

      // Cargar e incrustar la imagen del producto en máxima calidad
      if (p.image) {
        try {
          let imgBuffer: Buffer | null = null;
          let ext: "png" | "jpeg" | "gif" = "jpeg";

          const rawImg = String(p.image).trim();
          if (rawImg.startsWith("/api/storage/objects/") || rawImg.startsWith("/storage/objects/")) {
            const fileId = rawImg.split("/").pop() || "";
            const filePath = path.join(UPLOADS_DIR, fileId);
            const seedPath = path.join(SEED_UPLOADS_DIR, fileId);
            if (fs.existsSync(filePath)) {
              imgBuffer = fs.readFileSync(filePath);
            } else if (fs.existsSync(seedPath)) {
              imgBuffer = fs.readFileSync(seedPath);
            }
            if (fileId.toLowerCase().endsWith(".png")) ext = "png";
            else if (fileId.toLowerCase().endsWith(".gif")) ext = "gif";
            else ext = "jpeg";
          } else if (rawImg.startsWith("data:image/")) {
            const commaIdx = rawImg.indexOf(",");
            if (commaIdx !== -1) {
              const mimeMatch = rawImg.match(/^data:image\/([a-zA-Z0-9]+);base64,/);
              if (mimeMatch && mimeMatch[1] === "png") ext = "png";
              else if (mimeMatch && mimeMatch[1] === "gif") ext = "gif";
              else ext = "jpeg";
              imgBuffer = Buffer.from(rawImg.slice(commaIdx + 1), "base64");
            }
          } else if (rawImg.startsWith("http://") || rawImg.startsWith("https://")) {
            try {
              const fetchRes = await fetch(rawImg, { signal: AbortSignal.timeout(4000) });
              if (fetchRes.ok) {
                const arrayBuf = await fetchRes.arrayBuffer();
                imgBuffer = Buffer.from(arrayBuf);
                const ct = fetchRes.headers.get("content-type") || "";
                if (ct.includes("png") || rawImg.toLowerCase().endsWith(".png")) ext = "png";
                else if (ct.includes("gif") || rawImg.toLowerCase().endsWith(".gif")) ext = "gif";
                else ext = "jpeg";
              }
            } catch {}
          }

          if (imgBuffer && imgBuffer.length > 0) {
            const imageId = workbook.addImage({
              buffer: imgBuffer as any,
              extension: ext,
            });

            worksheet.addImage(imageId, {
              tl: { col: 6, row: excelRowIdx - 1 }, // 0-indexed: col 6 es Columna G (Imagen), row excelRowIdx-1
              ext: { width: 65, height: 65 },
              editAs: "oneCell",
            });
          }
        } catch (imgErr) {
          console.error(`Error attaching image for product ${p.name}:`, imgErr);
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const dateStr = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Disposition", `attachment; filename="productos_fellas_con_imagenes_${dateStr}.xlsx"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(Buffer.from(buffer as any));
  } catch (err: any) {
    console.error("Error exporting Excel with images:", err);
    res.status(500).json({ error: "Error al generar archivo Excel con imágenes: " + err.message });
  }
});

apiRouter.post("/admin/import-excel", async (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) {
      res.status(400).json({ error: "No se proporcionaron datos del archivo Excel" });
      return;
    }
    const buffer = Buffer.from(base64, "base64");

    // Extraer de forma asíncrona todas las imágenes adjuntas dentro del archivo XLSX (OpenXML DrawingML / Media)
    const { rowImages, cellImages, allMediaList, allMediaByName } = await extractImagesFromXlsx(buffer);

    const workbook = XLSX.read(buffer, { type: "buffer" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:Z1000");
    const headerRow = range.s.r; // Fila de cabeceras en Excel (típicamente 0)
    const rows: any[] = XLSX.utils.sheet_to_json(worksheet);

    let count = 0;
    let extractedCount = 0;
    const usedMediaFilenames = new Set<string>();

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const excelRow = headerRow + 1 + i; // Índice de fila 0-indexed en Excel
      const keys = Object.keys(row);
      const normalize = (s: string) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");

      const findKey = (candidates: string[]) => {
        for (const c of candidates) {
          const nc = normalize(c);
          for (const k of keys) {
            const nk = normalize(k);
            if (nk === nc || nk.includes(nc) || nc.includes(nk)) {
              return row[k];
            }
          }
        }
        return undefined;
      };

      const name = String(findKey(["nombre", "producto", "titulo", "name", "product", "articulo", "artículo"]) || "").trim();
      if (!name) continue;

      const description = String(findKey(["descripcion", "descripccion", "detalles", "det", "desc", "description"]) || "").trim();
      
      const rawPrice = findKey(["precio", "valor", "costo", "price", "$"]);
      let price = 0;
      if (typeof rawPrice === "number") {
        price = isNaN(rawPrice) ? 0 : rawPrice;
      } else if (rawPrice) {
        const digits = String(rawPrice).replace(/[^0-9]/g, "");
        price = digits ? parseInt(digits, 10) : 0;
      }

      const aisle = String(findKey(["pasillo", "seccion", "sección", "zona", "aisle", "departamento"]) || "General").trim();
      const category = String(findKey(["categoria", "categoría", "category", "familia"]) || "General").trim();
      const subcategory = String(findKey(["subcategoria", "subcategoría", "sub-categoria", "subcategory", "subfamilia"]) || "General").trim();
      
      const rawImageVal = findKey(["imagen", "foto", "image", "url", "img", "fotografia", "adjunto", "archivo"]);
      let finalImageUrl = "";

      // 1. Verificar si hay imagen adjunta vinculada a esta fila exacta del Excel
      if (rowImages[excelRow] && !usedMediaFilenames.has(rowImages[excelRow].filename)) {
        const media = rowImages[excelRow];
        const saved = saveExtractedImage(media.buffer, `${name}_${media.filename}`);
        finalImageUrl = saved.url;
        usedMediaFilenames.add(media.filename);
        extractedCount++;
      }

      // 2. Si no hay por anclaje directo de fila, verificar el texto de la celda de imagen
      if (!finalImageUrl && rawImageVal) {
        const rawStr = String(rawImageVal).trim();
        if (rawStr.startsWith("data:image/")) {
          // Data URI Base64
          const commaIdx = rawStr.indexOf(",");
          if (commaIdx !== -1) {
            const mimeMatch = rawStr.match(/^data:image\/([a-zA-Z0-9]+);base64,/);
            const ext = mimeMatch ? mimeMatch[1] : "jpg";
            const b64Data = rawStr.slice(commaIdx + 1);
            const imgBuf = Buffer.from(b64Data, "base64");
            const saved = saveExtractedImage(imgBuf, `${name}_excel.${ext}`, ext);
            finalImageUrl = saved.url;
            extractedCount++;
          }
        } else if (rawStr.startsWith("http://") || rawStr.startsWith("https://")) {
          // URL externa
          finalImageUrl = rawStr;
        } else if (allMediaByName[rawStr.toLowerCase()]) {
          // El texto de la celda coincide con el nombre de un archivo adjunto en el zip
          const media = allMediaByName[rawStr.toLowerCase()];
          const saved = saveExtractedImage(media.buffer, `${name}_${media.filename}`);
          finalImageUrl = saved.url;
          usedMediaFilenames.add(media.filename);
          extractedCount++;
        }
      }

      // 3. Fallback: Si la fila no tenía imagen pero el Excel traía imágenes adjuntas sin asignar
      if (!finalImageUrl) {
        const unusedMedia = allMediaList.find((m) => !usedMediaFilenames.has(m.filename));
        if (unusedMedia) {
          const saved = saveExtractedImage(unusedMedia.buffer, `${name}_${unusedMedia.filename}`);
          finalImageUrl = saved.url;
          usedMediaFilenames.add(unusedMedia.filename);
          extractedCount++;
        }
      }

      // 4. Si definitivamente no hay imagen adjunta ni texto
      if (!finalImageUrl) {
        finalImageUrl = "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=600&auto=format&fit=crop&q=80";
      }

      const ofertaRaw = findKey(["oferta", "en oferta", "oferta?", "promo", "promocion"]);
      const oferta = ofertaRaw === true || String(ofertaRaw).toLowerCase() === 'si' || String(ofertaRaw).toLowerCase() === 'true' || String(ofertaRaw) === '1' || String(ofertaRaw).toLowerCase() === 'sí';

      // Ensure aisle exists
      const existingAisles = dbManager.getAisles();
      if (!existingAisles.some(a => a.name.toLowerCase() === aisle.toLowerCase())) {
        dbManager.addAisle(aisle);
      }

      // Ensure category exists
      const existingCategories = dbManager.getCategories();
      if (!existingCategories.some(c => c.name.toLowerCase() === category.toLowerCase())) {
        dbManager.addCategory(category);
      }

      // Ensure subcategory exists
      const existingSubcategories = dbManager.getSubcategories();
      if (!existingSubcategories.some(s => s.name.toLowerCase() === subcategory.toLowerCase())) {
        dbManager.addSubcategory(subcategory);
      }

      dbManager.addProduct({
        name,
        description,
        price,
        image: finalImageUrl,
        category,
        aisle,
        subcategory,
        optionsTitle: "",
        options: [],
        oferta,
        depositoEnabled: false,
        depositoAmount: 0,
        position: dbManager.getProducts().length,
        publishedSocial: true,
        hidden: false,
        transferenciaEnabled: false,
        transferenciaAmount: 0,
      });
      count++;
    }

    res.json({ ok: true, importedCount: count, imagesExtracted: extractedCount });
  } catch (err: any) {
    console.error("Excel import error:", err);
    res.status(500).json({ error: "Error al procesar el archivo Excel: " + err.message });
  }
});

// Backup & Persistence routes for GitHub / Render deploys
apiRouter.get("/admin/backup/export", async (_req, res) => {
  try {
    const zip = new JSZip();

    // 1. Database
    const dbData = dbManager.getRaw();
    zip.file("fellas_db.json", JSON.stringify(dbData, null, 2));

    // 2. Media library metadata
    const mediaMeta = getMediaLibraryMeta();
    zip.file("media_library.json", JSON.stringify(mediaMeta, null, 2));

    // 3. Uploaded photos
    const uploadsFolder = zip.folder("uploads");
    if (fs.existsSync(UPLOADS_DIR) && uploadsFolder) {
      const files = fs.readdirSync(UPLOADS_DIR);
      for (const file of files) {
        const fullPath = path.join(UPLOADS_DIR, file);
        try {
          if (fs.statSync(fullPath).isFile()) {
            const fileData = fs.readFileSync(fullPath);
            uploadsFolder.file(file, fileData);
          }
        } catch {}
      }
    }

    const zipBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });

    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    res.setHeader("Content-Disposition", `attachment; filename="fellas_backup_${timestamp}.zip"`);
    res.setHeader("Content-Type", "application/zip");
    res.send(zipBuffer);
  } catch (err: any) {
    console.error("Backup export error:", err);
    res.status(500).json({ error: "Error al exportar copia de seguridad: " + err.message });
  }
});

apiRouter.get("/admin/backup/export-json", (_req, res) => {
  try {
    const dbData = dbManager.getRaw();
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 10);
    res.setHeader("Content-Disposition", `attachment; filename="fellas_catalogo_${timestamp}.json"`);
    res.setHeader("Content-Type", "application/json");
    res.send(JSON.stringify(dbData, null, 2));
  } catch (err: any) {
    res.status(500).json({ error: "Error al exportar JSON: " + err.message });
  }
});

apiRouter.post("/admin/backup/restore", async (req, res) => {
  try {
    const { base64, filename } = req.body;
    if (!base64) {
      res.status(400).json({ error: "No se proporcionaron datos de archivo para restaurar" });
      return;
    }

    let cleanBase64 = base64;
    if (cleanBase64.includes(";base64,")) {
      cleanBase64 = cleanBase64.split(";base64,")[1];
    }
    const buffer = Buffer.from(cleanBase64, "base64");

    const isZip = (filename && filename.toLowerCase().endsWith(".zip")) ||
      (buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b);

    if (isZip) {
      const zip = await JSZip.loadAsync(buffer);
      let productsCount = 0;
      let restoredImagesCount = 0;

      // 1. Restore fellas_db.json
      const dbEntry = zip.file("fellas_db.json") || Object.values(zip.files).find((f) => f.name.endsWith("fellas_db.json") || f.name.endsWith(".json"));
      if (dbEntry) {
        const text = await dbEntry.async("string");
        const parsed = JSON.parse(text);
        if (parsed) {
          dbManager.restoreRaw(parsed);
          productsCount = parsed.products?.length || 0;
          try {
            fs.writeFileSync(path.join(DATA_DIR, "fellas_db.json"), JSON.stringify(parsed, null, 2), "utf-8");
            fs.writeFileSync(path.resolve(process.cwd(), "src", "server", "seed_db.json"), JSON.stringify(parsed, null, 2), "utf-8");
          } catch {}
        }
      }

      // 2. Restore media_library.json
      const mediaEntry = zip.file("media_library.json") || Object.values(zip.files).find((f) => f.name.endsWith("media_library.json"));
      if (mediaEntry) {
        try {
          const mediaText = await mediaEntry.async("string");
          const mediaParsed = JSON.parse(mediaText);
          if (Array.isArray(mediaParsed)) {
            saveMediaLibraryMeta(mediaParsed);
          }
        } catch {}
      }

      // 3. Restore all uploaded files in uploads/
      const zipFiles = Object.keys(zip.files);
      for (const relPath of zipFiles) {
        const entry = zip.files[relPath];
        if (entry.dir) continue;
        if (relPath.startsWith("uploads/") || (!relPath.endsWith(".json") && !relPath.endsWith(".md"))) {
          const fileName = path.basename(relPath);
          if (!fileName || fileName.startsWith(".")) continue;
          const fileBuf = await entry.async("nodebuffer");
          const destUpload = path.join(UPLOADS_DIR, fileName);
          const destSeed = path.join(SEED_UPLOADS_DIR, fileName);
          fs.writeFileSync(destUpload, fileBuf);
          try {
            fs.writeFileSync(destSeed, fileBuf);
          } catch {}
          restoredImagesCount++;
        }
      }

      // Re-index media library
      const metaList = getMediaLibraryMeta();
      const metaMap = new Map<string, MediaMetaItem>(metaList.map((m) => [m.id, m]));
      if (fs.existsSync(UPLOADS_DIR)) {
        const diskFiles = fs.readdirSync(UPLOADS_DIR);
        for (const file of diskFiles) {
          if (!metaMap.has(file)) {
            try {
              const stats = fs.statSync(path.join(UPLOADS_DIR, file));
              if (stats.isFile()) {
                metaMap.set(file, {
                  id: file,
                  name: `Imagen_${file.slice(0, 8)}.jpg`,
                  url: `/api/storage/objects/${file}`,
                  createdAt: stats.mtime.toISOString(),
                  size: stats.size,
                });
              }
            } catch {}
          }
        }
        saveMediaLibraryMeta(Array.from(metaMap.values()));
      }

      res.json({
        ok: true,
        message: `¡Copia de seguridad restaurada exitosamente! ${productsCount} productos y ${restoredImagesCount} imágenes recuperadas.`,
        productsCount,
        restoredImagesCount,
      });
    } else {
      // JSON format
      const text = buffer.toString("utf-8");
      const parsed = JSON.parse(text);
      if (!parsed || (!parsed.products && !parsed.settings)) {
        res.status(400).json({ error: "El archivo JSON no tiene un formato válido de base de datos Fellas" });
        return;
      }
      dbManager.restoreRaw(parsed);
      const productsCount = parsed.products?.length || 0;
      try {
        fs.writeFileSync(path.join(DATA_DIR, "fellas_db.json"), JSON.stringify(parsed, null, 2), "utf-8");
        fs.writeFileSync(path.resolve(process.cwd(), "src", "server", "seed_db.json"), JSON.stringify(parsed, null, 2), "utf-8");
      } catch {}

      res.json({
        ok: true,
        message: `¡Base de datos restaurada! ${productsCount} productos recuperados.`,
        productsCount,
      });
    }
  } catch (err: any) {
    console.error("Backup restore error:", err);
    res.status(500).json({ error: "Error al restaurar copia de seguridad: " + err.message });
  }
});

apiRouter.post("/products/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderProducts(ids);
  res.json({ ok: true });
});

// 8. Orders
apiRouter.get("/orders", (_req, res) => {
  res.json(dbManager.getOrders());
});

apiRouter.post("/orders", (req, res) => {
  const body = req.body;
  const order = dbManager.addOrder({
    customerName: body.customerName,
    address: body.address || "",
    phone: body.phone || "",
    notes: body.notes || "",
    items: body.items || [],
    subtotal: body.subtotal ?? null,
    discountCode: body.discountCode ?? null,
    discountAmount: body.discountAmount ?? null,
    total: Number(body.total) || 0,
    status: body.status || "Pendiente",
    caja: body.caja ?? null,
    modificationNote: body.modificationNote ?? null,
  });
  res.status(201).json(order);
});

apiRouter.patch("/orders/:id", (req, res) => {
  const id = Number(req.params.id);
  const updated = dbManager.updateOrder(id, req.body);
  if (!updated) {
    res.status(404).json({ error: "Order not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/orders/:id", (req, res) => {
  const ok = dbManager.deleteOrder(Number(req.params.id));
  res.json({ ok });
});

// 9. Admin
apiRouter.post("/admin/login", (req, res) => {
  const password = String(req.body.password || "").trim();
  const configuredPassword = process.env.ADMIN_PASSWORD || "fellhonpm";
  const deliveryPassword = process.env.DELIVERY_PASSWORD || "botifelldely";

  if (password === "fellhonpm" || password === configuredPassword) {
    res.json({ ok: true, role: "admin" });
    return;
  }
  if (password === "botifelldely" || password === deliveryPassword) {
    res.json({ ok: true, role: "delivery" });
    return;
  }
  res.status(401).json({ ok: false, error: "Contraseña incorrecta" });
});

apiRouter.get("/admin/stats", (_req, res) => {
  res.json({
    productCount: dbManager.getProducts().length,
    categoryCount: dbManager.getCategories().length,
    aisleCount: dbManager.getAisles().length,
    subcategoryCount: dbManager.getSubcategories().length,
  });
});

apiRouter.get("/admin/visits/stats", (_req, res) => {
  const visits = dbManager.getVisits();
  const now = Date.now();
  const dayMs = 24 * 3600 * 1000;
  const startOfToday = new Date().setHours(0, 0, 0, 0);

  const todayCount = visits.filter((v) => new Date(v.createdAt).getTime() >= startOfToday).length;
  const last7Count = visits.filter((v) => new Date(v.createdAt).getTime() >= now - 7 * dayMs).length;
  const last30Count = visits.filter((v) => new Date(v.createdAt).getTime() >= now - 30 * dayMs).length;

  const hourMap = new Map<number, number>();
  for (let h = 0; h < 24; h++) hourMap.set(h, 0);

  const cityMap = new Map<string, { city: string; country: string; count: number }>();
  const countryMap = new Map<string, number>();
  const dayMap = new Map<string, number>();

  for (const v of visits) {
    const vDate = new Date(v.createdAt);
    if (vDate.getTime() >= now - 7 * dayMs) {
      const hour = vDate.getHours();
      hourMap.set(hour, (hourMap.get(hour) || 0) + 1);
    }
    if (vDate.getTime() >= now - 30 * dayMs) {
      if (v.city) {
        const k = `${v.city}_${v.country}`;
        const prev = cityMap.get(k) || { city: v.city, country: v.country, count: 0 };
        prev.count += 1;
        cityMap.set(k, prev);
      }
      if (v.country) {
        countryMap.set(v.country, (countryMap.get(v.country) || 0) + 1);
      }
      const dayStr = vDate.toISOString().slice(0, 10);
      dayMap.set(dayStr, (dayMap.get(dayStr) || 0) + 1);
    }
  }

  res.json({
    total: visits.length,
    today: todayCount,
    last7Days: last7Count,
    last30Days: last30Count,
    hourlyDistribution: Array.from(hourMap.entries()).map(([hour, count]) => ({ hour, count })),
    topCities: Array.from(cityMap.values()).sort((a, b) => b.count - a.count).slice(0, 10),
    topCountries: Array.from(countryMap.entries())
      .map(([country, count]) => ({ country, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10),
    dailySeries: Array.from(dayMap.entries())
      .map(([day, count]) => ({ day, count }))
      .sort((a, b) => a.day.localeCompare(b.day)),
  });
});

apiRouter.get("/admin/sales/stats", (req, res) => {
  const period = req.query.period === "monthly" ? "monthly" : "weekly";
  const days = period === "monthly" ? 30 : 7;
  const since = Date.now() - days * 24 * 3600 * 1000;
  const orders = dbManager.getOrders();

  const windowOrders = orders.filter((o) => new Date(o.createdAt).getTime() >= since);
  const delivered = windowOrders.filter((o) => o.status === "Entregado");
  const cancelled = windowOrders.filter((o) => o.status === "Anulado");

  const totalRevenue = delivered.reduce((acc, o) => acc + (Number(o.total) || 0), 0);

  const productStats = new Map<string, { qty: number; revenue: number }>();
  const hourMap = new Map<number, { orders: number; revenue: number }>();
  const dowMap = new Map<number, { orders: number; revenue: number }>();
  const dayMap = new Map<string, { orders: number; revenue: number }>();

  for (let h = 0; h < 24; h++) hourMap.set(h, { orders: 0, revenue: 0 });
  for (let d = 0; d < 7; d++) dowMap.set(d, { orders: 0, revenue: 0 });

  for (const o of delivered) {
    const oDate = new Date(o.createdAt);
    const hour = oDate.getHours();
    const dow = oDate.getDay();
    const day = oDate.toISOString().slice(0, 10);
    const rev = Number(o.total) || 0;

    const hEntry = hourMap.get(hour)!;
    hEntry.orders += 1;
    hEntry.revenue += rev;

    const dEntry = dowMap.get(dow)!;
    dEntry.orders += 1;
    dEntry.revenue += rev;

    const dayEntry = dayMap.get(day) || { orders: 0, revenue: 0 };
    dayEntry.orders += 1;
    dayEntry.revenue += rev;
    dayMap.set(day, dayEntry);

    for (const it of o.items || []) {
      const q = Number(it.quantity) || 1;
      const p = Number(it.price) || 0;
      const prev = productStats.get(it.name) || { qty: 0, revenue: 0 };
      prev.qty += q;
      prev.revenue += q * p;
      productStats.set(it.name, prev);
    }
  }

  const topProducts = Array.from(productStats.entries())
    .map(([name, data]) => ({ name, qty: data.qty, revenue: data.revenue }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 10);

  res.json({
    period,
    days,
    totalOrders: delivered.length,
    cancelledOrders: cancelled.length,
    totalRevenue,
    topProducts,
    hourlyDistribution: Array.from(hourMap.entries()).map(([hour, val]) => ({
      hour,
      orders: val.orders,
      revenue: val.revenue,
    })),
    dayOfWeekDistribution: Array.from(dowMap.entries()).map(([dayOfWeek, val]) => ({
      dayOfWeek,
      orders: val.orders,
      revenue: val.revenue,
    })),
    dailySeries: Array.from(dayMap.entries())
      .map(([day, val]) => ({ day, orders: val.orders, revenue: val.revenue }))
      .sort((a, b) => a.day.localeCompare(b.day)),
  });
});

apiRouter.get("/admin/customers", (_req, res) => {
  const orders = dbManager.getOrders();
  const customerMap = new Map<string, { customer_name: string; phone: string; order_count: number; total_spent: number }>();

  for (const o of orders) {
    if (o.status === "Anulado") continue;
    const phone = o.phone || "Sin teléfono";
    const prev = customerMap.get(phone) || {
      customer_name: o.customerName || "Cliente",
      phone,
      order_count: 0,
      total_spent: 0,
    };
    prev.order_count += 1;
    prev.total_spent += Number(o.total) || 0;
    customerMap.set(phone, prev);
  }

  const topBuyers = Array.from(customerMap.values())
    .sort((a, b) => b.total_spent - a.total_spent)
    .slice(0, 20)
    .map((c) => ({
      name: c.customer_name,
      phone: c.phone,
      orderCount: c.order_count,
      totalSpent: c.total_spent,
    }));

  res.json({
    registered: [],
    topBuyers,
  });
});

// 10. Visits
apiRouter.post("/visits", async (req, res) => {
  const ip = (req.headers["x-forwarded-for"] as string)?.split(",")[0]?.trim() || req.socket.remoteAddress || "";
  const geo = await lookupIpGeo(ip);
  const pathStr = String(req.body.path || "/");
  const userAgent = String(req.headers["user-agent"] || "");

  const visit = dbManager.addVisit({
    ipAddress: ip,
    city: geo.city,
    region: geo.region,
    country: geo.country,
    path: pathStr,
    userAgent,
  });

  res.json({ ok: true, id: visit.id });
});

// 11. Delivery locations
apiRouter.get("/delivery-locations", (_req, res) => {
  res.json(dbManager.getDeliveryLocations().filter((l) => l.active));
});

apiRouter.get("/admin/delivery-locations", (_req, res) => {
  res.json(dbManager.getDeliveryLocations());
});

apiRouter.post("/admin/delivery-locations", (req, res) => {
  const body = req.body;
  const created = dbManager.addDeliveryLocation({
    name: body.name,
    price: Number(body.price) || 0,
    active: body.active !== undefined ? Boolean(body.active) : true,
    position: dbManager.getDeliveryLocations().length,
  });
  res.status(201).json(created);
});

apiRouter.patch("/admin/delivery-locations/:id", (req, res) => {
  const updated = dbManager.updateDeliveryLocation(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Location not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/admin/delivery-locations/:id", (req, res) => {
  const ok = dbManager.deleteDeliveryLocation(Number(req.params.id));
  res.json({ ok });
});

// 12. Discounts
apiRouter.post("/discounts/apply", (req, res) => {
  const { code, orderTotal } = req.body;
  if (!code) {
    res.status(400).json({ error: "Código requerido" });
    return;
  }
  const discount = dbManager.getDiscountByCode(code);
  if (!discount) {
    res.status(404).json({ error: "Cupón no válido o inactivo" });
    return;
  }
  if (discount.minOrder && Number(orderTotal) < discount.minOrder) {
    res.status(400).json({
      error: `El pedido mínimo para este cupón es $${discount.minOrder.toLocaleString("es-CL")}`,
    });
    return;
  }
  let discountAmount = 0;
  if (discount.type === "percentage") {
    discountAmount = Math.round((Number(orderTotal) * discount.amount) / 100);
  } else {
    discountAmount = discount.amount;
  }
  res.json({
    valid: true,
    code: discount.code,
    type: discount.type,
    amount: discount.amount,
    discountAmount,
  });
});

apiRouter.get("/admin/discounts", (_req, res) => {
  res.json(dbManager.getDiscounts());
});

apiRouter.post("/admin/discounts", (req, res) => {
  const body = req.body;
  const created = dbManager.addDiscount({
    code: String(body.code).toUpperCase().trim(),
    type: body.type || "percentage",
    amount: Number(body.amount) || 0,
    customerId: body.customerId ?? null,
    active: body.active !== undefined ? Boolean(body.active) : true,
    minOrder: Number(body.minOrder) || 0,
    usesLeft: body.usesLeft !== undefined ? Number(body.usesLeft) : null,
    expiresAt: body.expiresAt ?? null,
  });
  res.status(201).json(created);
});

apiRouter.delete("/admin/discounts/:id", (req, res) => {
  const ok = dbManager.deleteDiscount(Number(req.params.id));
  res.json({ ok });
});

// 13. Customers
apiRouter.post("/customers/register", async (req, res) => {
  const { email, password, name, phone } = req.body;
  if (!email || !password) {
    res.status(400).json({ error: "Email y contraseña requeridos" });
    return;
  }
  const existing = dbManager.getCustomerByEmail(email);
  if (existing) {
    res.status(409).json({ error: "El email ya está registrado" });
    return;
  }
  const passwordHash = await bcrypt.hash(password, 10);
  const customer = dbManager.addCustomer({
    email,
    passwordHash,
    name: name || "",
    phone: phone || "",
  });
  const token = jwt.sign({ customerId: customer.id }, JWT_SECRET, { expiresIn: "30d" });
  res.status(201).json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone }, token });
});

apiRouter.post("/customers/login", async (req, res) => {
  const { email, password } = req.body;
  const customer = dbManager.getCustomerByEmail(email);
  if (!customer) {
    res.status(401).json({ error: "Credenciales inválidas" });
    return;
  }
  const valid = await bcrypt.compare(password, customer.passwordHash);
  if (!valid) {
    res.status(401).json({ error: "Credenciales inválidas" });
    return;
  }
  const token = jwt.sign({ customerId: customer.id }, JWT_SECRET, { expiresIn: "30d" });
  res.json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone }, token });
});

apiRouter.get("/customers/me", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const customer = dbManager.getCustomerById(customerId);
  if (!customer) {
    res.status(404).json({ error: "Cliente no encontrado" });
    return;
  }
  res.json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone } });
});

apiRouter.get("/customers/cart", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const cart = dbManager.getSavedCart(customerId);
  res.json({ cart: cart ? cart.items : [] });
});

apiRouter.post("/customers/cart", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const items = req.body.items || [];
  const saved = dbManager.saveCart(customerId, "Mi Carrito", items);
  res.json({ ok: true, cart: saved.items });
});

// 14. File Storage & Media Library for batch images
apiRouter.get("/media-library", (_req, res) => {
  try {
    const metaList = getMediaLibraryMeta();
    const metaMap = new Map<string, MediaMetaItem>(metaList.map((m) => [m.id, m]));
    
    // Auto-discover files on disk
    if (fs.existsSync(UPLOADS_DIR)) {
      const files = fs.readdirSync(UPLOADS_DIR);
      let changed = false;

      for (const file of files) {
        if (!metaMap.has(file)) {
          try {
            const stats = fs.statSync(path.join(UPLOADS_DIR, file));
            if (stats.isFile()) {
              const detected = classifyImageHeuristic(file);
              const newItem: MediaMetaItem = {
                id: file,
                name: `Imagen_${file.slice(0, 8)}.jpg`,
                url: `/api/storage/objects/${file}`,
                createdAt: stats.mtime.toISOString(),
                size: stats.size,
                category: detected.category,
                aiDetectedTitle: detected.title,
              };
              metaMap.set(file, newItem);
              changed = true;
            }
          } catch {}
        }
      }

      // Fill in category for any existing items that don't have one
      for (const item of metaMap.values()) {
        if (!item.category) {
          const detected = classifyImageHeuristic(item.name || item.id);
          item.category = detected.category;
          if (detected.title && !item.aiDetectedTitle) item.aiDetectedTitle = detected.title;
          changed = true;
        }
      }

      if (changed) {
        saveMediaLibraryMeta(Array.from(metaMap.values()));
      }
    }

    // Return all items sorted newest first without destructive purging
    const validItems: MediaMetaItem[] = Array.from(metaMap.values());
    validItems.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

    res.json({ images: validItems, categories: PRODUCT_IMAGE_SECTIONS });
  } catch (err) {
    console.error("Error reading media library:", err);
    res.status(500).json({ error: "Failed to load media library" });
  }
});

// Batch upload endpoint: accepts multiple images at once
apiRouter.post("/media-library/upload-batch", async (req, res) => {
  try {
    const { files } = req.body as { files: Array<{ name: string; data: string }> };
    if (!Array.isArray(files) || files.length === 0) {
      res.status(400).json({ error: "No files provided in files array" });
      return;
    }

    const currentMeta = getMediaLibraryMeta();
    const newItems: MediaMetaItem[] = [];

    // First write files and set heuristic default
    const savedFiles: Array<{ fileId: string; buffer: Buffer; name: string; mimeType: string }> = [];

    for (const item of files) {
      if (!item.data) continue;
      const fileId = randomUUID();
      const filePath = path.join(UPLOADS_DIR, fileId);

      // Extract base64
      let base64Data = item.data;
      let mimeType = "image/jpeg";
      if (base64Data.includes(";base64,")) {
        const parts = base64Data.split(";base64,");
        const mimeMatch = parts[0].match(/data:(.*?)$/);
        if (mimeMatch) mimeType = mimeMatch[1];
        base64Data = parts[1];
      }

      const buffer = Buffer.from(base64Data, "base64");
      fs.writeFileSync(filePath, buffer);
      try {
        fs.writeFileSync(path.join(SEED_UPLOADS_DIR, fileId), buffer);
      } catch {}

      savedFiles.push({ fileId, buffer, name: item.name || `Imagen_${fileId.slice(0, 8)}.jpg`, mimeType });
    }

    // Now classify using AI in parallel if GEMINI_API_KEY is present
    const apiKey = process.env.GEMINI_API_KEY;

    for (const saved of savedFiles) {
      let detected: { category: ProductImageSection; title?: string };

      if (apiKey && saved.buffer.length > 0) {
        try {
          detected = await classifyImageWithAI(saved.buffer, saved.name, saved.mimeType);
        } catch {
          detected = classifyImageHeuristic(saved.name || saved.fileId);
        }
      } else {
        detected = classifyImageHeuristic(saved.name || saved.fileId);
      }

      const metaItem: MediaMetaItem = {
        id: saved.fileId,
        name: saved.name,
        url: `/api/storage/objects/${saved.fileId}`,
        createdAt: new Date().toISOString(),
        size: saved.buffer.length,
        category: detected.category,
        aiDetectedTitle: detected.title,
      };

      newItems.push(metaItem);
      currentMeta.unshift(metaItem);
    }

    saveMediaLibraryMeta(currentMeta);
    res.json({ ok: true, uploaded: newItems, total: currentMeta.length });
  } catch (err) {
    console.error("Batch upload error:", err);
    res.status(500).json({ error: "Failed to upload batch images" });
  }
});

// Clasificación / Organización masiva con Inteligencia Artificial
apiRouter.post("/media-library/classify-ai", async (req, res) => {
  try {
    const { forceAll = false } = req.body || {};
    const currentMeta = getMediaLibraryMeta();
    let updatedCount = 0;

    for (const item of currentMeta) {
      if (forceAll || !item.category || item.category === "Otros / General") {
        let buffer: Buffer | null = null;
        const filePath = path.join(UPLOADS_DIR, item.id);
        const seedPath = path.join(SEED_UPLOADS_DIR, item.id);
        if (fs.existsSync(filePath)) {
          buffer = fs.readFileSync(filePath);
        } else if (fs.existsSync(seedPath)) {
          buffer = fs.readFileSync(seedPath);
        }

        if (buffer && buffer.length > 0) {
          const classified = await classifyImageWithAI(buffer, item.name || item.id);
          item.category = classified.category;
          if (classified.title) {
            item.aiDetectedTitle = classified.title;
          }
          updatedCount++;
        } else {
          const heuristic = classifyImageHeuristic(item.name || item.id);
          item.category = heuristic.category;
          if (heuristic.title) item.aiDetectedTitle = heuristic.title;
          updatedCount++;
        }
      }
    }

    saveMediaLibraryMeta(currentMeta);
    res.json({ ok: true, updatedCount, images: currentMeta, categories: PRODUCT_IMAGE_SECTIONS });
  } catch (err: any) {
    console.error("Error in AI classification of media library:", err);
    res.status(500).json({ error: "Error classifying media: " + err.message });
  }
});

// Actualizar categoría o nombre de una imagen
apiRouter.patch("/media-library/:id", (req, res) => {
  try {
    const fileId = req.params.id;
    const { category, name, aiDetectedTitle } = req.body;
    const currentMeta = getMediaLibraryMeta();
    const item = currentMeta.find((m) => m.id === fileId);
    if (!item) {
      res.status(404).json({ error: "Image not found" });
      return;
    }

    if (category !== undefined) item.category = category;
    if (name !== undefined) item.name = name;
    if (aiDetectedTitle !== undefined) item.aiDetectedTitle = aiDetectedTitle;

    saveMediaLibraryMeta(currentMeta);
    res.json({ ok: true, image: item });
  } catch (err: any) {
    res.status(500).json({ error: "Failed to update image" });
  }
});

apiRouter.delete("/media-library/all", (_req, res) => {
  try {
    let deletedCount = 0;
    if (fs.existsSync(UPLOADS_DIR)) {
      const files = fs.readdirSync(UPLOADS_DIR);
      for (const file of files) {
        try {
          fs.unlinkSync(path.join(UPLOADS_DIR, file));
          deletedCount++;
        } catch {}
      }
    }
    if (fs.existsSync(SEED_UPLOADS_DIR)) {
      const files = fs.readdirSync(SEED_UPLOADS_DIR);
      for (const file of files) {
        try {
          fs.unlinkSync(path.join(SEED_UPLOADS_DIR, file));
          deletedCount++;
        } catch {}
      }
    }
    saveMediaLibraryMeta([]);
    
    // Eliminar también las referencias de fotos de TODOS los productos
    const productsCleared = dbManager.clearAllProductImages();

    res.json({ ok: true, deletedCount, productsCleared });
  } catch (err: any) {
    console.error("Delete all media error:", err);
    res.status(500).json({ error: "Failed to delete all images: " + err.message });
  }
});

apiRouter.delete("/media-library/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  try {
    const metaList = getMediaLibraryMeta();
    const targetMeta = metaList.find((m) => m.id === fileId);

    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
    const seedPath = path.join(SEED_UPLOADS_DIR, fileId);
    if (fs.existsSync(seedPath)) {
      try {
        fs.unlinkSync(seedPath);
      } catch {}
    }
    const currentMeta = metaList.filter((m) => m.id !== fileId);
    saveMediaLibraryMeta(currentMeta);

    // Eliminar la foto de los productos que tengan asignada esta imagen
    let productsCleared = dbManager.clearProductImageByIdOrUrl(fileId);
    if (targetMeta?.name) {
      productsCleared += dbManager.clearProductImageByIdOrUrl(targetMeta.name);
    }
    if (targetMeta?.url) {
      productsCleared += dbManager.clearProductImageByIdOrUrl(targetMeta.url);
    }

    res.json({ ok: true, productsCleared });
  } catch (err) {
    console.error("Delete media error:", err);
    res.status(500).json({ error: "Failed to delete image" });
  }
});

apiRouter.post("/storage/uploads/request-url", (req, res) => {
  const fileId = randomUUID();
  const fileName = req.body?.name || `Foto_${fileId.slice(0, 8)}.jpg`;
  try {
    const currentMeta = getMediaLibraryMeta();
    if (!currentMeta.some((m) => m.id === fileId)) {
      currentMeta.unshift({
        id: fileId,
        name: fileName,
        url: `/api/storage/objects/${fileId}`,
        createdAt: new Date().toISOString(),
      });
      saveMediaLibraryMeta(currentMeta);
    }
  } catch {}
  res.json({
    uploadURL: `/api/storage/uploads/${fileId}`,
    objectPath: `/objects/${fileId}`,
    fileId,
  });
});

apiRouter.put("/storage/uploads/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  const writeStream = fs.createWriteStream(filePath);
  req.pipe(writeStream);
  writeStream.on("finish", () => {
    try {
      const stats = fs.statSync(filePath);
      // Auto-mirror to seed uploads so fresh deploys / Git commits preserve the file
      try {
        fs.copyFileSync(filePath, path.join(SEED_UPLOADS_DIR, fileId));
      } catch {}
      const currentMeta = getMediaLibraryMeta();
      const existing = currentMeta.find((m) => m.id === fileId);
      if (existing) {
        existing.size = stats.size;
      } else {
        currentMeta.unshift({
          id: fileId,
          name: `Foto_${fileId.slice(0, 8)}.jpg`,
          url: `/api/storage/objects/${fileId}`,
          createdAt: new Date().toISOString(),
          size: stats.size,
        });
      }
      saveMediaLibraryMeta(currentMeta);
    } catch {}
    res.status(200).json({ ok: true });
  });
  writeStream.on("error", (err) => {
    console.error("Upload error:", err);
    res.status(500).json({ error: "Upload failed" });
  });
});

apiRouter.get("/storage/objects/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  if (fs.existsSync(filePath)) {
    res.sendFile(filePath);
  } else {
    // Check seed uploads as fallback
    const seedPath = path.join(SEED_UPLOADS_DIR, fileId);
    if (fs.existsSync(seedPath)) {
      try {
        fs.copyFileSync(seedPath, filePath);
      } catch {}
      res.sendFile(seedPath);
    } else {
      res.status(404).json({ error: "Object not found" });
    }
  }
});

apiRouter.delete("/storage/objects/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  if (fs.existsSync(filePath)) {
    try {
      fs.unlinkSync(filePath);
    } catch {}
  }
  const seedPath = path.join(SEED_UPLOADS_DIR, fileId);
  if (fs.existsSync(seedPath)) {
    try {
      fs.unlinkSync(seedPath);
    } catch {}
  }
  const currentMeta = getMediaLibraryMeta().filter((m) => m.id !== fileId);
  saveMediaLibraryMeta(currentMeta);
  res.json({ ok: true });
});

// 14b. Full Backup, Restore & Seed Engine (Prevents any data loss across GitHub & Render deploys)
apiRouter.get("/backup/download", (_req, res) => {
  try {
    const dbData = dbManager.getRaw();
    const mediaList = getMediaLibraryMeta();
    const uploadsMap: Record<string, string> = {};

    // Collect all uploaded images in base64
    const filesToRead = new Set<string>();
    if (fs.existsSync(UPLOADS_DIR)) {
      for (const f of fs.readdirSync(UPLOADS_DIR)) filesToRead.add(f);
    }
    if (fs.existsSync(SEED_UPLOADS_DIR)) {
      for (const f of fs.readdirSync(SEED_UPLOADS_DIR)) filesToRead.add(f);
    }

    for (const file of filesToRead) {
      try {
        const p = fs.existsSync(path.join(UPLOADS_DIR, file))
          ? path.join(UPLOADS_DIR, file)
          : path.join(SEED_UPLOADS_DIR, file);
        if (fs.statSync(p).isFile()) {
          const buf = fs.readFileSync(p);
          uploadsMap[file] = buf.toString("base64");
        }
      } catch {}
    }

    const backupPayload = {
      app: "fellas_market",
      version: 1,
      exportedAt: new Date().toISOString(),
      stats: {
        productsCount: dbData.products?.length || 0,
        categoriesCount: dbData.categories?.length || 0,
        imagesCount: Object.keys(uploadsMap).length,
      },
      database: dbData,
      mediaLibrary: mediaList,
      uploads: uploadsMap,
    };

    const dateStr = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Disposition", `attachment; filename="fellas_market_backup_${dateStr}.json"`);
    res.setHeader("Content-Type", "application/json");
    res.send(JSON.stringify(backupPayload, null, 2));
  } catch (err) {
    console.error("Error creating full backup:", err);
    res.status(500).json({ error: "Failed to generate backup" });
  }
});

apiRouter.post("/backup/restore", (req, res) => {
  try {
    const { database, mediaLibrary, uploads } = req.body;
    if (!database || !Array.isArray(database.products)) {
      res.status(400).json({ error: "Formato de respaldo inválido: faltan los datos del catálogo." });
      return;
    }

    // 1. Restore Database
    dbManager.restoreRaw(database);

    // 2. Restore Media Library
    if (Array.isArray(mediaLibrary)) {
      saveMediaLibraryMeta(mediaLibrary);
    }

    // 3. Unpack all uploaded image files
    let restoredImages = 0;
    if (uploads && typeof uploads === "object") {
      for (const [fileId, base64Data] of Object.entries(uploads)) {
        if (typeof base64Data === "string" && base64Data.length > 0) {
          try {
            let clean = base64Data;
            if (clean.includes(";base64,")) {
              clean = clean.split(";base64,")[1];
            }
            const buf = Buffer.from(clean, "base64");
            fs.writeFileSync(path.join(UPLOADS_DIR, fileId), buf);
            try {
              fs.writeFileSync(path.join(SEED_UPLOADS_DIR, fileId), buf);
            } catch {}
            restoredImages++;
          } catch {}
        }
      }
    }

    res.json({
      ok: true,
      message: "¡Copia de seguridad restaurada con éxito!",
      productsCount: database.products.length,
      imagesCount: restoredImages,
    });
  } catch (err) {
    console.error("Error restoring backup:", err);
    res.status(500).json({ error: "Error al restaurar el respaldo" });
  }
});

apiRouter.post("/backup/bake-seed", (_req, res) => {
  try {
    // 1. Bake database
    const dbData = dbManager.getRaw();
    const seedDbPath = path.resolve(process.cwd(), "src", "server", "seed_db.json");
    fs.writeFileSync(seedDbPath, JSON.stringify(dbData, null, 2), "utf-8");

    // 2. Bake media library
    const mediaList = getMediaLibraryMeta();
    const seedMediaPath = path.resolve(process.cwd(), "src", "server", "seed_media.json");
    fs.writeFileSync(seedMediaPath, JSON.stringify(mediaList, null, 2), "utf-8");

    // 3. Bake uploads into seed_uploads
    let count = 0;
    if (fs.existsSync(UPLOADS_DIR)) {
      for (const file of fs.readdirSync(UPLOADS_DIR)) {
        const src = path.join(UPLOADS_DIR, file);
        const dest = path.join(SEED_UPLOADS_DIR, file);
        if (fs.statSync(src).isFile()) {
          try {
            fs.copyFileSync(src, dest);
            count++;
          } catch {}
        }
      }
    }

    res.json({
      ok: true,
      message: "¡Datos fijados con éxito en el código fuente para GitHub y Render!",
      productsCount: dbData.products?.length || 0,
      imagesCount: count,
    });
  } catch (err) {
    console.error("Error baking seed:", err);
    res.status(500).json({ error: "Error al fijar datos" });
  }
});

// 15. Resolve OpenGraph image
apiRouter.get("/resolve-image", async (req, res) => {
  const url = req.query.url as string;
  if (!url) {
    res.status(400).json({ error: "url param required" });
    return;
  }
  try {
    const response = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0" },
      signal: AbortSignal.timeout(8000),
    });
    const html = await response.text();
    const ogMatch =
      html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i) ||
      html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);
    if (ogMatch?.[1]) {
      res.json({ resolved: ogMatch[1] });
      return;
    }
    res.status(422).json({ error: "No direct image found" });
  } catch {
    res.status(502).json({ error: "Failed to fetch URL" });
  }
});

// 16. AI Copywriting with Gemini
apiRouter.post("/ai/improve-banner", async (req, res) => {
  const { title = "", subtitle = "" } = req.body;
  if (!title.trim() && !subtitle.trim()) {
    res.status(400).json({ error: "Empty input" });
    return;
  }

  const prompt = `Eres un copywriter experto en marketing juvenil para una botillería chilena llamada "Fella's Market" en Alerce, Puerto Montt.
Mejora el siguiente título y subtítulo de un banner promocional:
Título actual: "${title}"
Subtítulo actual: "${subtitle}"

Reglas:
- Tono juvenil, llamativo, fiestero y con buena onda chilena.
- TÍTULO: máximo 4-6 palabras, impactante, en MAYÚSCULAS con 1-2 emojis.
- SUBTÍTULO: máximo 10-14 palabras, frase atractiva con hashtag corto (#Alerce, #FindeSemana, #Promos).
- Responde estrictamente en formato JSON: {"title": "...", "subtitle": "..."}`;

  const apiKey = process.env.GEMINI_API_KEY;
  if (apiKey) {
    try {
      const ai = new GoogleGenAI({ apiKey });
      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: prompt,
        config: { responseMimeType: "application/json" },
      });
      const parsed = JSON.parse(response.text || "{}");
      if (parsed.title && parsed.subtitle) {
        res.json({ title: parsed.title, subtitle: parsed.subtitle });
        return;
      }
    } catch (err) {
      console.warn("Gemini call fallback:", err);
    }
  }

  // High quality fallback
  res.json({
    title: (title || "FELLA'S MARKET EN VIVO 🍻").toUpperCase(),
    subtitle: (subtitle || "Las mejores promos y delivery a la puerta de tu casa #Alerce"),
  });
});
