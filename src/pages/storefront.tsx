import {
  useState,
  useMemo,
  useEffect,
  useRef,
  Fragment,
  FormEvent,
  ChangeEvent,
  DragEvent,
} from "react";
import { motion, AnimatePresence } from "motion/react";

import { useQueryClient } from "@tanstack/react-query";
import {
  useGetMenu,
  useUpdateSettings,
  useCreateProduct,
  useUpdateProduct,
  useDeleteProduct,
  useCreateCategory,
  useDeleteCategory,
  useReorderCategories,
  useCreateAisle,
  useDeleteAisle,
  useUpdateAisle,
  useCreateSubcategory,
  useDeleteSubcategory,
  useAdminLogin,
  useListOrders,
  useCreateOrder,
  useDeleteOrder,
  useUpdateOrderStatus,
  useTrackVisit,
  useGetVisitStats,
  useGetSalesStats,
  useReorderAisles,
  useReorderSubcategories,
  useReorderProducts,
  getGetMenuQueryKey,
  getListOrdersQueryKey,
  Product,
  Settings,
  Category,
  Order,
  BannerSlide,
  DeliveryLocation,
  useListDeliveryLocations,
  useCreateDeliveryLocation,
  useUpdateDeliveryLocation,
  useDeleteDeliveryLocation,
  getListDeliveryLocationsQueryKey,
} from "@workspace/api-client-react";
import { generateStatsPDF } from "../lib/statsPdf";
import { MediaLibraryModal, type MediaItem } from "../components/media-library-modal";
import { MediaAdminPanel } from "../components/media-admin-panel";

declare module "@workspace/api-client-react" {
  interface Settings {
    topAnnouncementText?: string;
    contingencyMode?: boolean;
    contingencyMessage?: string;
    contingencyBannerImage?: string;
    contingencyAislesConfig?: Record<string, { enabled?: boolean; bannerImage?: string; noticeText?: string; productIds?: number[] }>;
    promoBannerImage?: string;
    aislesBannerImage?: string;
    homeCollectionProductIds?: number[];
    recommendedProductIds?: number[];
    opportunitiesBannerImage?: string;
    opportunitiesBannerTitle?: string;
    opportunitiesBannerSubtitle?: string;
    opportunitiesProductIds?: number[];
    packsBannerImage?: string;
    packsBannerTitle?: string;
    packsBannerSubtitle?: string;
    packsProductIds?: number[];
    footerLogo?: string;
    footerDescription?: string;
    socialInstagram?: string;
    socialFacebook?: string;
    socialTiktok?: string;
    socialWhatsapp?: string;
    googleMapsPlaceId?: string;
    googleMapsRating?: number;
    googleMapsReviewsCount?: number;
    googleMapsReviews?: Array<{
      id: string;
      author_name: string;
      author_photo?: string;
      rating: number;
      relative_time_description: string;
      text: string;
    }>;
    galleryTitle?: string;
    gallerySubtitle?: string;
    galleryImages?: Array<{
      id: string;
      url: string;
      title?: string;
      aisle?: string;
      buttonText?: string;
      caption?: string;
      link?: string;
    }>;
  }
  interface Product {
    contingencyEnabled?: boolean;
    description?: string;
  }
  interface ProductInput {
    contingencyEnabled?: boolean;
  }
}
import {
  ShoppingCart,
  Settings as SettingsIcon,
  Plus,
  Trash2,
  Phone,
  LogOut,
  X,
  Check,
  Utensils,
  Edit3,
  Upload,
  Search,
  GripVertical,
  CheckCircle,
  Layers,
  Tag,
  MapPin,
  ClipboardList,
  User,
  BarChart3,
  TrendingUp,
  Globe,
  Clock,
  Calendar,
  Eye,
  EyeOff,
  Flame,
  ChevronDown,
  ChevronRight,
  Download,
  Instagram,
  Facebook,
  MessageCircle,
  Video,
  ArrowUp,
  ArrowDown,
  ChevronUp,
  Minus,
  Sparkles,
  MapPinIcon,
  PhoneCall,
  ClockIcon,
  Link2,
  HardDriveUpload,
  Ban,
  Pencil,
  CreditCard,
  Share2,
  FolderOpen,
  ImageIcon,
  Copy,
  UserPlus,
  Percent,
  Bookmark,
  KeyRound,
  Gift,
  Users,
  Loader2,
  ShoppingBag,
  FileSpreadsheet,
  Menu,
  Package,
  Filter,
  LayoutGrid,
  ChevronLeft,
  ArrowLeft,
  Star,
  LogIn,
  AlertTriangle,
  ExternalLink,
  Camera,
  ShieldCheck,
  Archive,
  DownloadCloud,
  UploadCloud,
  Database,
} from "lucide-react";
import * as XLSX from "xlsx";
import { StoreGallery, DEFAULT_GALLERY_IMAGES, GalleryImageItem } from "@/components/store-gallery";

function downloadProductsExcel(products: Product[]) {
  const sorted = [...products].sort((a, b) => {
    const cat = (a.category ?? "").localeCompare(b.category ?? "", "es");
    if (cat !== 0) return cat;
    const aisle = (a.aisle ?? "").localeCompare(b.aisle ?? "", "es");
    if (aisle !== 0) return aisle;
    const sub = (a.subcategory ?? "").localeCompare(b.subcategory ?? "", "es");
    if (sub !== 0) return sub;
    return (a.name ?? "").localeCompare(b.name ?? "", "es");
  });

  const headers = ["codigo", "nombre", "variante", "categoria", "proveedor", "precio", "costo", "stock", "stock_minimo"];
  const rows = sorted.map(p => [
    p.id ?? "",
    p.name ?? "",
    p.subcategory ?? "",
    p.category ?? "",
    p.aisle ?? "",
    p.price ?? 0,
    "",
    "",
    "",
  ]);

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  ws["!cols"] = [
    { wch: 10 }, { wch: 30 }, { wch: 25 }, { wch: 18 }, { wch: 20 },
    { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 12 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Productos");
  XLSX.writeFile(wb, "plantilla_productos_POS.xlsx");
}

const subcategoryEmoji = (name: string): string => {
  const n = name.toLowerCase();
  if (/cerveza|chela|beer|lager|ipa|stout/.test(n)) return "🍺";
  if (/champ|espumant|sparkl|prosec/.test(n)) return "🍾";
  if (/vino|tinto|blanco|rosado|cabernet|merlot|carmen/.test(n)) return "🍷";
  if (/pisco|whisky|whiskey|ron|vodka|tequila|gin|destil|licor/.test(n)) return "🥃";
  if (/cocktail|coctel|trago|mix|piscola|fernet/.test(n)) return "🍹";
  if (/sake|asia/.test(n)) return "🍶";
  if (/snack|papa|chips|frito|maní|mani|nuez|fruto/.test(n)) return "🍿";
  if (/dulce|chocolate|caramelo|gomit|galleta/.test(n)) return "🍫";
  if (/agua|jugo|gaseos|bebid|soda|cola/.test(n)) return "🥤";
  if (/energ|red bull|monster/.test(n)) return "⚡";
  if (/cigarr|tabac|vape/.test(n)) return "🚬";
  if (/hielo|cold|frio/.test(n)) return "🧊";
  if (/promo|oferta|combo/.test(n)) return "🎉";
  if (/import/.test(n)) return "🌎";
  if (/nacional|chile/.test(n)) return "🇨🇱";
  return "✨";
};

const DEFAULT_AISLE_BANNERS: Record<string, string> = {
  cervezas: "https://images.unsplash.com/photo-1608270199929-e0930fca6a15?w=1600&auto=format&fit=crop&q=80",
  destilados: "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=1600&auto=format&fit=crop&q=80",
  vinos: "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=1600&auto=format&fit=crop&q=80",
  snacks: "https://images.unsplash.com/photo-1621939514649-280e2ee25f60?w=1600&auto=format&fit=crop&q=80",
  bebidas: "https://images.unsplash.com/photo-1622483767028-3f66f32aef97?w=1600&auto=format&fit=crop&q=80",
  packs: "https://images.unsplash.com/photo-1543007630-9710e4a00a20?w=1600&auto=format&fit=crop&q=80",
  promociones: "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=1600&auto=format&fit=crop&q=80",
  cigarrillos: "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=1600&auto=format&fit=crop&q=80",
  hielo: "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=1600&auto=format&fit=crop&q=80",
  dulces: "https://images.unsplash.com/photo-1582293041079-7814c2f12063?w=1600&auto=format&fit=crop&q=80",
};

function getAisleBannerImage(aisleName?: string, customImage?: string): string {
  if (customImage && customImage.trim().length > 0) return customImage;
  if (!aisleName) return "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=1600&auto=format&fit=crop&q=80";
  const nameNorm = aisleName.toLowerCase();
  if (/cerveza|beer|chela|lager|ipa|stout|pilsen/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.cervezas;
  if (/destilad|licor|whisky|whiskey|pisco|ron|vodka|gin|tequila|coctel|cocktail/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.destilados;
  if (/vino|wine|tinto|blanco|rosado|espumante|champagne|prosec/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.vinos;
  if (/snack|papa|chips|maní|mani|nuez|fruto|papas|frito/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.snacks;
  if (/bebida|jugo|gaseosa|soda|agua|energ/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.bebidas;
  if (/pack|combo|promo|oferta/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.packs;
  if (/cigarro|cigarrillo|tabaco|vape|vapeador/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.cigarrillos;
  if (/hielo|ice/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.hielo;
  if (/dulce|chocolat|galleta|caramelo/.test(nameNorm)) return DEFAULT_AISLE_BANNERS.dulces;
  return "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=1600&auto=format&fit=crop&q=80";
}

const normalize = (s: string) =>
  s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

function levenshtein(a: string, b: string): number {
  if (Math.abs(a.length - b.length) > 3) return 99;
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const prev = Array.from({ length: n + 1 }, (_, i) => i);
  const curr = new Array<number>(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      curr[j] = a[i - 1] === b[j - 1]
        ? prev[j - 1]
        : 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
    }
    for (let j = 0; j <= n; j++) prev[j] = curr[j];
  }
  return curr[n];
}

function fuzzyMatch(query: string, target: string): boolean {
  if (!query) return true;
  const q = normalize(query);
  const t = normalize(target);
  // Exact substring match first
  if (t.includes(q)) return true;
  // Each query word must match some target word
  const qWords = q.split(/\s+/).filter(Boolean);
  const tWords = t.split(/\s+/).filter(Boolean);
  return qWords.every((qw) =>
    tWords.some((tw) => {
      // Target word starts with what the user typed (prefix match)
      if (tw.startsWith(qw)) return true;
      // Only allow Levenshtein for words of 4+ chars, and be strict
      if (qw.length < 4 || tw.length < 4) return false;
      const maxDist = qw.length >= 7 ? 2 : 1;
      return levenshtein(qw, tw) <= maxDist;
    }),
  );
}

interface CartItem extends Product {
  quantity: number;
  selectedOption: string | null;
  cartItemId: number;
}

const DEFAULT_FORM = {
  name: "",
  price: "",
  image: "",
  category: "",
  aisle: "",
  subcategory: "",
  optionsTitle: "",
  optionsString: "",
  oferta: false,
  depositoEnabled: false,
  depositoAmount: "500",
  transferenciaEnabled: false,
  transferenciaAmount: "0",
  contingencyEnabled: false,
};

function useDragScroll() {
  const ref = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    let isDown = false;
    let moved = false;
    let startX = 0;
    let startScroll = 0;
    const onMouseDown = (e: MouseEvent) => {
      isDown = true;
      moved = false;
      startX = e.pageX;
      startScroll = el.scrollLeft;
      el.style.cursor = "grabbing";
    };
    const onMouseUp = () => {
      if (!isDown) return;
      isDown = false;
      el.style.cursor = "grab";
    };
    const onMouseMove = (e: MouseEvent) => {
      if (!isDown) return;
      const dx = e.pageX - startX;
      if (Math.abs(dx) > 5) {
        moved = true;
        el.scrollLeft = startScroll - dx;
        e.preventDefault();
      }
    };
    const onClickCapture = (e: MouseEvent) => {
      if (moved) {
        e.stopPropagation();
        e.preventDefault();
        moved = false;
      }
    };
    const onWheel = (e: WheelEvent) => {
      if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
        el.scrollLeft += e.deltaY;
      }
    };
    el.style.cursor = "grab";
    el.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mouseup", onMouseUp);
    window.addEventListener("mousemove", onMouseMove);
    el.addEventListener("click", onClickCapture, true);
    el.addEventListener("wheel", onWheel, { passive: true });
    return () => {
      el.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("mousemove", onMouseMove);
      el.removeEventListener("click", onClickCapture, true);
      el.removeEventListener("wheel", onWheel);
    };
  }, []);
  return ref;
}

export default function Storefront() {
  const queryClient = useQueryClient();
  const { data: menu, isLoading } = useGetMenu({
    query: {
      queryKey: getGetMenuQueryKey(),
      refetchInterval: 5000, // Poll every 5 seconds to automatically keep the storefront and admin in sync
    },
  });
  const { data: deliveryLocations = [] } = useListDeliveryLocations();

  const [ageVerified, setAgeVerified] = useState<boolean>(() => {
    try { return sessionStorage.getItem("fella_age_ok") === "1"; } catch { return false; }
  });
  const handleAgeYes = () => {
    try { sessionStorage.setItem("fella_age_ok", "1"); } catch {}
    setAgeVerified(true);
  };
  const handleAgeNo = () => {
    window.location.replace("https://www.google.com");
  };

  const [showScrollTop, setShowScrollTop] = useState(false);
  useEffect(() => {
    const onScroll = () => setShowScrollTop(window.scrollY > 400);
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  const [keyboardOffset, setKeyboardOffset] = useState(0);
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    const update = () => {
      const offset = window.innerHeight - (vv.offsetTop + vv.height);
      setKeyboardOffset(Math.max(0, offset));
    };
    vv.addEventListener("resize", update);
    vv.addEventListener("scroll", update);
    return () => {
      vv.removeEventListener("resize", update);
      vv.removeEventListener("scroll", update);
    };
  }, []);
  const scrollToTop = () => window.scrollTo({ top: 0, behavior: "smooth" });

  const settings: Settings = menu?.settings ?? {
    logo: "",
    pageTitle: "Urban Bite",
    favicon: "",
    bannerText: "",
    bannerDescription: "",
    bannerImage: "",
    bannerSlides: [],
    whatsapp: "",
    agencyName: "",
    agencyLogo: "",
    announcements: [],
    contactPhone: "",
    contactAddress: "",
    contactHours: "",
    openTime: "11:00",
    closeTime: "23:00",
    deliveryMinimum: 10000,
  };
  const categoriesData: Category[] = menu?.categories ?? [];
  const aislesData = menu?.aisles ?? [];
  const subcategoriesData = menu?.subcategories ?? [];
  const products: Product[] = menu?.products ?? [];

  const categories = categoriesData.map((c) => c.name);
  const aisles = aislesData.map((a) => a.name);
  const subcategories = subcategoriesData.map((s) => s.name);

  const updateSettingsMut = useUpdateSettings();
  const createProductMut = useCreateProduct();
  const updateProductMut = useUpdateProduct();
  const deleteProductMut = useDeleteProduct();
  const createCategoryMut = useCreateCategory();
  const deleteCategoryMut = useDeleteCategory();
  const reorderCategoriesMut = useReorderCategories();
  const createAisleMut = useCreateAisle();
  const deleteAisleMut = useDeleteAisle();
  const updateAisleMut = useUpdateAisle();
  const reorderAislesMut = useReorderAisles();
  const [aisleBannerDrafts, setAisleBannerDrafts] = useState<Record<string, string>>({});
  const createSubcategoryMut = useCreateSubcategory();
  const deleteSubcategoryMut = useDeleteSubcategory();
  const reorderSubcategoriesMut = useReorderSubcategories();
  const reorderProductsMut = useReorderProducts();
  const adminLoginMut = useAdminLogin();
  const [googleMapsData, setGoogleMapsData] = useState<{
    rating: number;
    reviewsCount: number;
    reviews: Array<{
      id: string;
      author_name: string;
      author_photo?: string;
      rating: number;
      relative_time_description: string;
      text: string;
    }>;
  }>({
    rating: 4.9,
    reviewsCount: 142,
    reviews: [
      {
        id: "g-1",
        author_name: "Carlos Soto",
        author_photo: "https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?auto=format&fit=crop&q=80&w=120",
        rating: 5,
        relative_time_description: "Hace 2 días",
        text: "¡Excelente atención y las cervezas siempre llegan ultra heladas! El delivery es súper rápido en Alerce. 100% recomendado.",
      },
      {
        id: "g-2",
        author_name: "Valentina Muñoz",
        author_photo: "https://images.unsplash.com/photo-1494790108377-be9c29b29330?auto=format&fit=crop&q=80&w=120",
        rating: 5,
        relative_time_description: "Hace una semana",
        text: "Salvaron nuestra junta de amigos un domingo a medianoche. Tienen de todo y el hielo nunca falta. ¡Geniales!",
      },
      {
        id: "g-3",
        author_name: "Matías Alarcón",
        author_photo: "https://images.unsplash.com/photo-1570295999919-56ceb5ecca61?auto=format&fit=crop&q=80&w=120",
        rating: 5,
        relative_time_description: "Hace 2 semanas",
        text: "Muy buena variedad de destilados y snacks. Los precios son justos y la página web es súper fácil de usar.",
      },
      {
        id: "g-4",
        author_name: "Camila Fernández",
        author_photo: "https://images.unsplash.com/photo-1544005313-94ddf0286df2?auto=format&fit=crop&q=80&w=120",
        rating: 5,
        relative_time_description: "Hace 3 semanas",
        text: "Pedimos por WhatsApp y la entrega llegó en 20 minutos exacta. Todo muy bien empaquetado y los tragos heladísimos. Se pasaron.",
      },
    ],
  });

  useEffect(() => {
    fetch("/api/google-maps-reviews")
      .then((res) => res.json())
      .then((data) => {
        if (data) {
          setGoogleMapsData(data);
        }
      })
      .catch(() => {});
  }, []);

  const refreshMenu = () =>
    queryClient.invalidateQueries({ queryKey: getGetMenuQueryKey() });

  const moveProduct = (productId: number, direction: "up" | "down", groupIds: number[]) => {
    const idx = groupIds.indexOf(productId);
    if (idx === -1) return;
    const newIdx = direction === "up" ? idx - 1 : idx + 1;
    if (newIdx < 0 || newIdx >= groupIds.length) return;
    const newGroup = [...groupIds];
    [newGroup[idx], newGroup[newIdx]] = [newGroup[newIdx], newGroup[idx]];
    const groupSet = new Set(groupIds);
    const allIds = products.map((p) => p.id);
    const newAllIds: number[] = [];
    let inserted = false;
    for (const id of allIds) {
      if (!groupSet.has(id)) {
        newAllIds.push(id);
      } else if (!inserted) {
        newAllIds.push(...newGroup);
        inserted = true;
      }
    }
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? {
        ...old,
        products: newAllIds.map((id) => old.products.find((p: any) => p.id === id)).filter(Boolean),
      } : old,
    );
    reorderProductsMut.mutate({ data: { ids: newAllIds } });
  };

  const [view, setView] = useState<"client" | "admin-login" | "admin">("client");
  const [cart, setCart] = useState<CartItem[]>([]);
  const [isCartOpen, setIsCartOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchQuery), 180);
    return () => clearTimeout(t);
  }, [searchQuery]);
  const [toast, setToast] = useState({
    show: false,
    message: "",
    actionLabel: "",
    onAction: null as null | (() => void),
  });
  const [activeCategory, setActiveCategory] = useState("");
  const [activeAisle, setActiveAisle] = useState("");
  const [showAisleMenu, setShowAisleMenu] = useState(false);
  const [showAuthBar, setShowAuthBar] = useState(false);
  const [showLocationBar, setShowLocationBar] = useState(false);
  const [navQuickFilter, setNavQuickFilter] = useState<"" | "oportunidades" | "packs">("");
  const [showComunasModal, setShowComunasModal] = useState(false);
  const [selectedComuna, setSelectedComuna] = useState("Santiago");
  const [activeAnnouncementIdx, setActiveAnnouncementIdx] = useState(0);
  const [showDedicatedProductsPage, setShowDedicatedProductsPage] = useState(false);
  const [dedicatedViewMode, setDedicatedViewMode] = useState<"catalog" | "oportunidades" | "packs" | "pasillo">("catalog");
  const [dedicatedSearchQuery, setDedicatedSearchQuery] = useState("");
  const [dedicatedSortBy, setDedicatedSortBy] = useState<"default" | "name_asc" | "name_desc" | "price_asc" | "price_desc">("default");
  const [dedicatedSubcatFilter, setDedicatedSubcatFilter] = useState<string>("all");

  useEffect(() => {
    if (searchQuery.trim().length > 0) {
      setShowDedicatedProductsPage(false);
    }
  }, [searchQuery]);

  const headerAnnouncements = useMemo(() => {
    if (settings.announcements && Array.isArray(settings.announcements) && settings.announcements.length > 0) {
      const valid = (settings.announcements as string[]).filter(
        (a) => a && typeof a === "string" && a.trim().length > 0
      );
      if (valid.length > 0) return valid;
    }
    return [
      "⚡ Envíos express y seguros directamente a tu puerta en minutos",
      "🔥 Promociones exclusivas y ofertas especiales todos los días",
      "💬 Atención directa y pedidos en línea vía WhatsApp 24/7",
      "📦 Gran variedad garantizada en todos nuestros pasillos",
    ];
  }, [settings.announcements]);

  useEffect(() => {
    if (headerAnnouncements.length <= 1) return;
    const interval = setInterval(() => {
      setActiveAnnouncementIdx((curr) => (curr + 1) % headerAnnouncements.length);
    }, 4500);
    return () => clearInterval(interval);
  }, [headerAnnouncements.length]);

  const [mobileNavDrawerOpen, setMobileNavDrawerOpen] = useState(false);
  const [checkoutOpen, setCheckoutOpen] = useState(false);
  const [checkoutStep, setCheckoutStep] = useState(1);
  const [checkoutForm, setCheckoutForm] = useState({
    customerName: "",
    address: "",
    phone: "",
    notes: "",
    deliveryType: "retiro" as "delivery" | "retiro",
  });

  const [optionModalInfo, setOptionModalInfo] = useState<
    { product: Product; selectedOption: string } | null
  >(null);
  const [adminTab, setAdminTab] =
    useState<"products" | "media" | "classifications" | "orders" | "stats" | "settings" | "social" | "customers" | "contingency">("products");
  const [mediaPickerOpen, setMediaPickerOpen] = useState(false);
  const [quickMediaImages, setQuickMediaImages] = useState<MediaItem[]>([]);

  const loadQuickMediaImages = async () => {
    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/media-library`);
      if (res.ok) {
        const data = await res.json();
        setQuickMediaImages(data.images || []);
      }
    } catch (err) {
      console.error("Error loading quick media library:", err);
    }
  };

  useEffect(() => {
    if (adminTab === "products" || adminTab === "media") {
      loadQuickMediaImages();
    }
  }, [adminTab]);
  const [adminRole, setAdminRole] = useState<"full" | "delivery" | null>(null);
  const [adminMobileMenuOpen, setAdminMobileMenuOpen] = useState(false);
  const [customerToken, setCustomerToken] = useState<string | null>(() => {
    try { return localStorage.getItem("customer_token"); } catch { return null; }
  });
  const [customer, setCustomer] = useState<{ id: number; email: string; name: string; phone: string } | null>(null);
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [authMode, setAuthMode] = useState<"login" | "register">("login");
  const [showAccountPanel, setShowAccountPanel] = useState(false);
  const [savedCarts, setSavedCarts] = useState<Array<{ id: number; name: string; items: any[]; createdAt: string }>>([]);
  const [selectedLocation, setSelectedLocation] = useState<DeliveryLocation | null>(null);
  const [discountCode, setDiscountCode] = useState("");
  const [appliedDiscount, setAppliedDiscount] = useState<{ code: string; type: string; amount: number; discountAmount: number; label: string } | null>(null);
  const [discountError, setDiscountError] = useState("");
  const [applyingDiscount, setApplyingDiscount] = useState(false);
  const [authFields, setAuthFields] = useState({ email: "", password: "", name: "", phone: "" });
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);

  const trackVisitMut = useTrackVisit();

  // Keep-alive ping para mantener el servidor activo en planes gratuitos (Render / Railway / Fly)
  useEffect(() => {
    const pingServer = () => {
      fetch("/api/ping").catch(() => {});
    };
    pingServer();
    const interval = setInterval(pingServer, 4 * 60 * 1000); // cada 4 minutos
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    try {
      const today = new Date().toISOString().slice(0, 10);
      const key = `ub_visit_${today}`;
      if (typeof window !== "undefined" && !window.sessionStorage.getItem(key)) {
        window.sessionStorage.setItem(key, "1");
        trackVisitMut.mutate({ data: { path: window.location.pathname || "/" } });
      }
    } catch {
      // ignore
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!customerToken) { setCustomer(null); return; }
    fetch("/api/customers/me", { headers: { Authorization: `Bearer ${customerToken}` } })
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (data) setCustomer(data);
        else {
          try { localStorage.removeItem("customer_token"); } catch {}
          setCustomerToken(null);
        }
      }).catch(() => {});
  }, [customerToken]);

  useEffect(() => {
    if (!customerToken || !customer) return;
    fetch("/api/customers/carts", { headers: { Authorization: `Bearer ${customerToken}` } })
      .then(r => r.ok ? r.json() : [])
      .then(setSavedCarts)
      .catch(() => {});
  }, [customer, customerToken]);

  const [editingProduct, setEditingProduct] = useState<number | null>(null);
  const [collapsedAisles, setCollapsedAisles] = useState<Record<string, boolean>>({});
  const [adminCategory, setAdminCategory] = useState<string>("Todas");
  const [adminProductSearch, setAdminProductSearch] = useState<string>("");
  const [contingencyProductSearch, setContingencyProductSearch] = useState<string>("");
  const [aisleSearchQueries, setAisleSearchQueries] = useState<Record<string, string>>({});
  const [contingencyStatusFilter, setContingencyStatusFilter] = useState<"all" | "active" | "inactive">("all");

  const toggleProductContingency = (prod: Product) => {
    const nextVal = prod.contingencyEnabled === false;
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old
        ? {
            ...old,
            products: old.products.map((p: any) =>
              p.id === prod.id ? { ...p, contingencyEnabled: nextVal } : p
            ),
          }
        : old
    );
    updateProductMut.mutate({
      id: prod.id,
      data: { ...prod, contingencyEnabled: nextVal },
    });
  };

  const batchProductContingency = (prods: Product[], enable: boolean) => {
    const ids = prods.map((p) => p.id);
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old
        ? {
            ...old,
            products: old.products.map((p: any) =>
              ids.includes(p.id) ? { ...p, contingencyEnabled: enable } : p
            ),
          }
        : old
    );
    fetch("/api/admin/products/contingency-batch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, enable }),
    }).catch(() => {});
  };

  const buyerCategoriesRef = useDragScroll();
  const buyerAislesRef = useDragScroll();
  const adminTabsRef = useDragScroll();
  const adminCategoriesRef = useDragScroll();

  const [newCategoryName, setNewCategoryName] = useState("");
  const [newAisleName, setNewAisleName] = useState("");
  const [newSubcatName, setNewSubcatName] = useState("");
  const [savedCats, setSavedCats] = useState(false);
  const [savedAisles, setSavedAisles] = useState(false);
  const [savedSubcats, setSavedSubcats] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");

  const [formState, setFormState] = useState(DEFAULT_FORM);
  const [settingsDraft, setSettingsDraft] = useState<Settings>(settings);
  const [activeSlide, setActiveSlide] = useState(0);
  const [showMobileSearch, setShowMobileSearch] = useState(false);

  // Admin Progressive Disclosure & Guided Steps State
  const [settingsSubTab, setSettingsSubTab] = useState<
    "all" | "brand" | "ticker" | "banners" | "featured" | "delivery" | "hours" | "contingency" | "contact" | "gallery" | "backup"
  >("all");
  const [openSettingsSections, setOpenSettingsSections] = useState<Record<string, boolean>>({
    brand: true,
    ticker: false,
    banners: false,
    featured: false,
    delivery: false,
    hours: false,
    contingency: false,
    contact: false,
    gallery: false,
    backup: false,
  });
  const [featuredRecoSearch, setFeaturedRecoSearch] = useState("");
  const [featuredCollecSearch, setFeaturedCollecSearch] = useState("");
  const [featuredOportunidadesSearch, setFeaturedOportunidadesSearch] = useState("");
  const [featuredPacksSearch, setFeaturedPacksSearch] = useState("");
  const toggleSettingsSection = (sec: string) => {
    setOpenSettingsSections((p) => ({ ...p, [sec]: !p[sec] }));
  };
  const [showAddProductModal, setShowAddProductModal] = useState<boolean>(false);
  const [productWizardStep, setProductWizardStep] = useState<1 | 2 | 3>(1);
  const [productViewMode, setProductViewMode] = useState<"grid" | "table">("grid");
  const [openProductFormSections, setOpenProductFormSections] = useState<{ main: boolean; media: boolean; extras: boolean }>({ main: true, media: false, extras: false });
  const [productFormSubTab, setProductFormSubTab] = useState<"all" | "main" | "media" | "extras">("main");
  const [productFilterSpecial, setProductFilterSpecial] = useState<"all" | "oferta" | "contingency" | "retornable">("all");
  const [classificationsSubTab, setClassificationsSubTab] = useState<"categories" | "aisles" | "subcategories" | "all">("categories");
  const [openClassificationsSections, setOpenClassificationsSections] = useState<Record<string, boolean>>({
    categories: true,
    aisles: true,
    subcategories: true,
  });
  const [openAislesAdmin, setOpenAislesAdmin] = useState<Record<number, boolean>>({});

  useEffect(() => {
    if (menu?.settings) setSettingsDraft(menu.settings);
  }, [menu?.settings]);

  const bannerSlides: BannerSlide[] = settings.bannerSlides ?? [];
  const effectiveSlides: BannerSlide[] =
    bannerSlides.length > 0
      ? bannerSlides
      : settings.bannerImage || settings.bannerText
        ? [{ image: settings.bannerImage, title: settings.bannerText, description: settings.bannerDescription }]
        : [];

  useEffect(() => {
    if (effectiveSlides.length <= 1) return;
    const timer = setInterval(() => {
      setActiveSlide((prev) => (prev + 1) % effectiveSlides.length);
    }, 4500);
    return () => clearInterval(timer);
  }, [effectiveSlides.length]);

  useEffect(() => {
    if (categories.length > 0 && (!activeCategory || !categories.includes(activeCategory))) {
      setActiveCategory(categories[0]);
    }
  }, [categories, activeCategory]);

  useEffect(() => {
    if (!menu) return;
    if (settings.pageTitle) document.title = settings.pageTitle;
    const iconUrl = settings.favicon || settings.logo;
    if (iconUrl) {
      let link = document.querySelector("link[rel~='icon']") as HTMLLinkElement | null;
      if (!link) {
        link = document.createElement("link");
        link.rel = "icon";
        document.head.appendChild(link);
      }
      link.href = iconUrl;
    }
  }, [menu, settings.pageTitle, settings.favicon, settings.logo]);

  const allStoreAisles = useMemo(() => {
    const fromData = aislesData.map((a) => a.name);
    const fromProds = Array.from(new Set(products.map((p) => p.aisle))).filter(Boolean);
    return Array.from(new Set([...fromData, ...fromProds]));
  }, [aislesData, products]);

  const promoProducts = useMemo(() => {
    const samplePromos: Product[] = [
      {
        id: 9001,
        name: "Pack Piscola Mistral 35° + Coca-Cola 1.5L + Hielo",
        price: 13990,
        category: "Promociones",
        aisle: "Packs y Promos",
        subcategory: "Packs",
        image: "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80",
        oferta: true,
        bestseller: true,
        options: [],
        optionsTitle: "",
        depositoEnabled: false,
        depositoAmount: 0,
        transferenciaEnabled: false,
        transferenciaAmount: 0,
        contingencyEnabled: false,
        hidden: false,
        publishedSocial: false,
        position: 0,
      },
      {
        id: 9002,
        name: "Combo Cerveza Corona Extra 6x330cc + Limones",
        price: 9490,
        category: "Cervezas",
        aisle: "Cervezas e Importadas",
        subcategory: "Promos",
        image: "https://images.unsplash.com/photo-1608270199929-e0930fca6a15?w=600&auto=format&fit=crop&q=80",
        oferta: true,
        bestseller: true,
        options: [],
        optionsTitle: "",
        depositoEnabled: false,
        depositoAmount: 0,
        transferenciaEnabled: false,
        transferenciaAmount: 0,
        contingencyEnabled: false,
        hidden: false,
        publishedSocial: false,
        position: 1,
      },
      {
        id: 9003,
        name: "Pack Fernet Branca 750ml + Coca-Cola 1.5L",
        price: 16990,
        category: "Licores",
        aisle: "Destilados y Licores",
        subcategory: "Packs",
        image: "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=600&auto=format&fit=crop&q=80",
        oferta: true,
        bestseller: false,
        options: [],
        optionsTitle: "",
        depositoEnabled: false,
        depositoAmount: 0,
        transferenciaEnabled: false,
        transferenciaAmount: 0,
        contingencyEnabled: false,
        hidden: false,
        publishedSocial: false,
        position: 2,
      },
      {
        id: 9004,
        name: "Promo Gin Bombay Sapphire 750cc + 4 Tónicas",
        price: 21990,
        category: "Gin",
        aisle: "Destilados y Licores",
        subcategory: "Packs",
        image: "https://images.unsplash.com/photo-1551538827-9c037cb4f32a?w=600&auto=format&fit=crop&q=80",
        oferta: true,
        bestseller: true,
        options: [],
        optionsTitle: "",
        depositoEnabled: false,
        depositoAmount: 0,
        transferenciaEnabled: false,
        transferenciaAmount: 0,
        contingencyEnabled: false,
        hidden: false,
        publishedSocial: false,
        position: 3,
      },
    ];

    if (!products || products.length === 0) return samplePromos;

    const customIds = settings.recommendedProductIds ?? [];
    if (customIds.length > 0) {
      const selected = customIds
        .map((id) => products.find((p) => p.id === id && !p.hidden))
        .filter((p): p is Product => Boolean(p));
      if (selected.length >= 4) return selected.slice(0, 4);

      const otherCandidates = products.filter(
        (p) =>
          !p.hidden &&
          !selected.some((s) => s.id === p.id) &&
          (p.oferta ||
            p.bestseller ||
            p.category?.toLowerCase().includes("pack") ||
            p.category?.toLowerCase().includes("promo") ||
            p.name?.toLowerCase().includes("pack") ||
            p.name?.toLowerCase().includes("promo") ||
            Boolean(p.image))
      );
      const combined = [...selected, ...otherCandidates];
      if (combined.length >= 4) return combined.slice(0, 4);
      const regular = products.filter((p) => !p.hidden && !combined.some((c) => c.id === p.id));
      return [...combined, ...regular, ...samplePromos].slice(0, 4);
    }

    const candidates = products.filter(
      (p) =>
        !p.hidden &&
        (p.oferta ||
          p.bestseller ||
          p.category?.toLowerCase().includes("pack") ||
          p.category?.toLowerCase().includes("promo") ||
          p.name?.toLowerCase().includes("pack") ||
          p.name?.toLowerCase().includes("promo") ||
          Boolean(p.image))
    );
    if (candidates.length >= 4) return candidates.slice(0, 4);
    const regular = products.filter((p) => !p.hidden && !candidates.some((c) => c.id === p.id));
    const combined = [...candidates, ...regular];
    if (combined.length >= 4) return combined.slice(0, 4);
    return [...combined, ...samplePromos].slice(0, 4);
  }, [products, settings.recommendedProductIds]);

  const baseProducts = useMemo(() => {
    if (!settings.contingencyMode) return products;
    const contingencyAisles = settings.contingencyAislesConfig || {};
    const filtered = products.filter((p) => {
      if (p.hidden) return false;
      const aisleKey = p.aisle || p.category;
      if (contingencyAisles[aisleKey] && contingencyAisles[aisleKey].enabled === false) {
        return false;
      }
      return p.contingencyEnabled !== false;
    });
    if (filtered.length > 0) return filtered;
    return products.filter((p) => !p.hidden);
  }, [products, settings.contingencyMode, settings.contingencyAislesConfig]);

  const allMenuSections = useMemo(() => {
    const list: Array<{ name: string; type: "category" | "aisle" }> = [];
    categories.forEach((c) => list.push({ name: c, type: "category" }));
    allStoreAisles.forEach((a) => {
      if (!list.some((item) => item.name.toLowerCase() === a.toLowerCase())) {
        list.push({ name: a, type: "aisle" });
      }
    });

    // Always filter out empty categories or aisles to avoid rendering dead/empty menu links
    return list.filter((item) => {
      const count = baseProducts.filter((p) =>
        item.type === "category" ? p.category === item.name : p.aisle === item.name
      ).length;
      return count > 0;
    });
  }, [categories, allStoreAisles, baseProducts]);

  const filteredProducts = useMemo(() => {
    if (debouncedSearch && debouncedSearch.trim().length > 0) {
      const q = normalize(debouncedSearch.trim());
      const terms = q.split(/\s+/).filter(Boolean);
      return baseProducts.filter((p) => {
        if (p.hidden) return false;
        const nameNorm = normalize(p.name || "");
        const catNorm = normalize(p.category || "");
        const aisleNorm = normalize(p.aisle || "");
        const subNorm = normalize(p.subcategory || "");

        return terms.every((term) =>
          nameNorm.includes(term) ||
          catNorm.includes(term) ||
          aisleNorm.includes(term) ||
          subNorm.includes(term) ||
          fuzzyMatch(term, p.name)
        );
      });
    }
    if (navQuickFilter === "oportunidades") {
      return baseProducts.filter((p) => !p.hidden && (p.oferta || p.bestseller));
    }
    if (navQuickFilter === "packs") {
      return baseProducts.filter(
        (p) =>
          !p.hidden &&
          (p.category?.toLowerCase().includes("pack") ||
            p.subcategory?.toLowerCase().includes("pack") ||
            p.aisle?.toLowerCase().includes("pack") ||
            p.name.toLowerCase().includes("pack") ||
            p.name.toLowerCase().includes("combo"))
      );
    }
    if (activeAisle) {
      return baseProducts.filter((p) => !p.hidden && p.aisle === activeAisle);
    }
    if (activeCategory) {
      return baseProducts.filter((p) => !p.hidden && p.category === activeCategory);
    }
    return baseProducts.filter((p) => !p.hidden);
  }, [baseProducts, activeCategory, activeAisle, debouncedSearch, navQuickFilter]);

  // Búsqueda instantánea y limpia: lista de productos que coinciden directamente con la búsqueda
  const searchResults = useMemo(() => {
    const q = searchQuery.trim();
    if (!q) return [];
    const normQ = normalize(q);
    const terms = normQ.split(/\s+/).filter(Boolean);
    return baseProducts.filter((p) => {
      if (p.hidden) return false;
      const nameNorm = normalize(p.name || "");
      const catNorm = normalize(p.category || "");
      const aisleNorm = normalize(p.aisle || "");
      const subNorm = normalize(p.subcategory || "");
      const descNorm = normalize(p.description || "");

      return terms.every((term) =>
        nameNorm.includes(term) ||
        catNorm.includes(term) ||
        aisleNorm.includes(term) ||
        subNorm.includes(term) ||
        descNorm.includes(term) ||
        fuzzyMatch(term, p.name)
      );
    });
  }, [searchQuery, baseProducts]);

  const homeCollectionProducts = useMemo(() => {
    const customIds = settings.homeCollectionProductIds ?? [];
    if (customIds.length > 0) {
      const selected = customIds
        .map((id) => baseProducts.find((p) => p.id === id && !p.hidden))
        .filter((p): p is Product => Boolean(p));
      if (selected.length >= 6) return selected.slice(0, 6);
      const remainder = baseProducts.filter(
        (p) => !p.hidden && !selected.some((s) => s.id === p.id)
      );
      const combined = [...selected, ...remainder];
      if (combined.length > 0) return combined.slice(0, 6);
    }
    const nonHidden = baseProducts.filter((p) => !p.hidden);
    return nonHidden.slice(0, 6);
  }, [baseProducts, settings.homeCollectionProductIds]);

  const opportunitiesCustomProducts = useMemo(() => {
    const customIds = settings.opportunitiesProductIds ?? [];
    if (customIds.length > 0) {
      const selected = customIds
        .map((id) => baseProducts.find((p) => p.id === id && !p.hidden))
        .filter((p): p is Product => Boolean(p));
      if (selected.length > 0) return selected;
    }
    const candidates = baseProducts.filter(
      (p) => !p.hidden && (p.oferta || p.bestseller)
    );
    if (candidates.length > 0) return candidates;
    return baseProducts.filter((p) => !p.hidden).slice(0, 12);
  }, [baseProducts, settings.opportunitiesProductIds]);

  const packsCustomProducts = useMemo(() => {
    const customIds = settings.packsProductIds ?? [];
    if (customIds.length > 0) {
      const selected = customIds
        .map((id) => baseProducts.find((p) => p.id === id && !p.hidden))
        .filter((p): p is Product => Boolean(p));
      if (selected.length > 0) return selected;
    }
    const candidates = baseProducts.filter(
      (p) =>
        !p.hidden &&
        (p.category?.toLowerCase().includes("pack") ||
          p.subcategory?.toLowerCase().includes("pack") ||
          p.aisle?.toLowerCase().includes("pack") ||
          p.name.toLowerCase().includes("pack") ||
          p.name.toLowerCase().includes("combo") ||
          p.name.toLowerCase().includes("promocion"))
    );
    if (candidates.length > 0) return candidates;
    return baseProducts.filter((p) => !p.hidden).slice(0, 12);
  }, [baseProducts, settings.packsProductIds]);

  const groupedByAisle = useMemo(
    () =>
      filteredProducts.reduce<Record<string, Product[]>>((acc, product) => {
        if (!acc[product.aisle]) acc[product.aisle] = [];
        acc[product.aisle].push(product);
        return acc;
      }, {}),
    [filteredProducts],
  );

  const activeAisles = useMemo(
    () => {
      const keys = Object.keys(groupedByAisle);
      let list: string[] = [];
      if (activeAisle) {
        list = [activeAisle].filter((a) => (groupedByAisle[a] || []).length > 0 || keys.includes(a));
      } else if (activeCategory) {
        list = keys.filter((k) => (groupedByAisle[k] || []).length > 0);
      } else {
        list = allStoreAisles.filter((name) => keys.includes(name)).concat(keys.filter((k) => !allStoreAisles.includes(k)));
      }

      // Strictly preserve custom administrative position order
      return [...list].sort((a, b) => {
        const idxA = allStoreAisles.indexOf(a);
        const idxB = allStoreAisles.indexOf(b);
        if (idxA === -1 && idxB === -1) return 0;
        if (idxA === -1) return 1;
        if (idxB === -1) return -1;
        return idxA - idxB;
      });
    },
    [allStoreAisles, groupedByAisle, activeAisle, activeCategory],
  );

  const cartTotal = useMemo(() => cart.reduce((sum, item) => {
    const fee = (item.transferenciaEnabled && item.transferenciaAmount) ? item.transferenciaAmount : 0;
    return sum + (item.price + fee) * item.quantity;
  }, 0), [cart]);
  const cartItemCount = useMemo(() => cart.reduce((count, item) => count + item.quantity, 0), [cart]);

  // Helper para parsear cualquier formato de hora (24h "11:00", 12h con AM/PM "11:00 AM", "23:45 PM", etc.) a minutos del día
  const parseTimeToMinutes = (rawStr?: string, defaultMinutes = 0): number => {
    if (!rawStr || typeof rawStr !== "string") return defaultMinutes;
    const str = rawStr.trim().toUpperCase();
    if (!str) return defaultMinutes;

    const isPM = str.includes("PM");
    const isAM = str.includes("AM");

    // Extraer solo números y dos puntos
    const cleaned = str.replace(/[^0-9:]/g, "");
    const parts = cleaned.split(":");
    let hours = parseInt(parts[0] || "0", 10);
    const minutes = parseInt(parts[1] || "0", 10);

    if (isNaN(hours)) hours = 0;
    const validMinutes = isNaN(minutes) ? 0 : Math.min(59, Math.max(0, minutes));

    if (isPM && hours < 12) {
      hours += 12;
    } else if (isAM && hours === 12) {
      hours = 0;
    }

    hours = Math.min(23, Math.max(0, hours));
    return hours * 60 + validMinutes;
  };

  // Verificación del estado de apertura de la tienda (robusto a AM/PM, 24h y zona horaria Chile / Local)
  const isStoreOpen = useMemo(() => {
    const openTime = settings.openTime || "11:00";
    const closeTime = settings.closeTime || "23:45";

    const openMinutes = parseTimeToMinutes(openTime, 11 * 60);
    const closeMinutes = parseTimeToMinutes(closeTime, 23 * 60 + 45);

    const now = new Date();
    const localMinutes = now.getHours() * 60 + now.getMinutes();

    // Obtener minutos actuales en zona horaria de Chile (America/Santiago) de forma robusta
    let chileMinutes = localMinutes;
    try {
      const formatted = new Intl.DateTimeFormat("es-CL", {
        timeZone: "America/Santiago",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(now);
      const parts = formatted.split(":");
      const h = parseInt(parts[0], 10) % 24;
      const m = parseInt(parts[1], 10);
      if (!isNaN(h) && !isNaN(m)) {
        chileMinutes = h * 60 + m;
      }
    } catch {
      chileMinutes = localMinutes;
    }

    const checkWindow = (curr: number): boolean => {
      // 24 horas abierto si apertura y cierre son iguales
      if (openMinutes === closeMinutes) return true;
      // Horario que cruza la medianoche (ej: 11:00 a 02:00 de la madrugada)
      if (closeMinutes < openMinutes) {
        return curr >= openMinutes || curr <= closeMinutes;
      }
      // Horario normal dentro del mismo día (ej: 11:00 a 23:45)
      return curr >= openMinutes && curr <= closeMinutes;
    };

    return checkWindow(chileMinutes) || checkWindow(localMinutes);
  }, [settings.openTime, settings.closeTime]);

  const deliveryMinimum = settings.deliveryMinimum ?? 10000;
  const hasLocations = deliveryLocations.length > 0;
  const DELIVERY_COST = selectedLocation?.price ?? (hasLocations ? 0 : 3000);
  const isDelivery = checkoutForm.deliveryType === "delivery";
  const orderTotal = cartTotal + (isDelivery ? DELIVERY_COST : 0);
  const discountSavings = appliedDiscount?.discountAmount ?? 0;
  const finalTotal = cartTotal - discountSavings + (isDelivery ? DELIVERY_COST : 0);
  const deliveryMinimumMet =
    !isDelivery || cartTotal >= deliveryMinimum;

  const showToast = (message: string) => {
    setToast({ show: true, message, actionLabel: "", onAction: null });
    setTimeout(() => setToast({ show: false, message: "", actionLabel: "", onAction: null }), 3000);
  };

  const applyDiscount = async () => {
    if (!discountCode.trim()) return;
    setApplyingDiscount(true);
    setDiscountError("");
    try {
      const r = await fetch("/api/discounts/apply", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(customerToken ? { Authorization: `Bearer ${customerToken}` } : {}) },
        body: JSON.stringify({ code: discountCode.trim(), orderTotal: cartTotal }),
      });
      const data = await r.json();
      if (!r.ok) { setDiscountError(data.message ?? "Código inválido."); return; }
      setAppliedDiscount(data);
    } catch { setDiscountError("Error al validar el código."); }
    finally { setApplyingDiscount(false); }
  };

  const saveCurrentCart = async () => {
    if (!customerToken || cart.length === 0) return;
    const name = `Carrito ${new Date().toLocaleDateString("es-CL")}`;
    try {
      const r = await fetch("/api/customers/carts", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${customerToken}` },
        body: JSON.stringify({ name, items: cart }),
      });
      if (r.ok) {
        const saved = await r.json();
        setSavedCarts(prev => [...prev, saved]);
        showToast("Carrito guardado ✓");
      }
    } catch {}
  };

  const loadSavedCart = (savedCart: { items: any[] }) => {
    setCart(savedCart.items.map((item: any, i: number) => ({ ...item, cartItemId: Date.now() + i })));
    setShowAccountPanel(false);
    showToast("Carrito cargado");
  };

  const deleteSavedCart = async (cartId: number) => {
    if (!customerToken) return;
    try {
      await fetch(`/api/customers/carts/${cartId}`, { method: "DELETE", headers: { Authorization: `Bearer ${customerToken}` } });
      setSavedCarts(prev => prev.filter(c => c.id !== cartId));
    } catch {}
  };

  const handleCustomerAuth = async (e: FormEvent) => {
    e.preventDefault();
    setAuthLoading(true);
    setAuthError("");

    const inputUser = authFields.email.trim().toLowerCase();
    const inputPass = authFields.password.trim();

    // Check Admin login credentials (admin / fellhonpm)
    if (authMode === "login" && (inputUser === "admin" || inputUser === "admin@fellas.cl" || inputUser === "admin@gmail.com") && inputPass === "fellhonpm") {
      setAdminRole("full");
      setView("admin");
      setShowAuthModal(false);
      setAuthFields({ email: "", password: "", name: "", phone: "" });
      showToast("Acceso concedido a Panel Administrador ✓");
      setAuthLoading(false);
      return;
    }

    // Check Delivery login credentials (delivery / botifelldely)
    if (authMode === "login" && (inputUser === "delivery" || inputUser === "delivery@fellas.cl" || inputUser === "delivery@gmail.com") && inputPass === "botifelldely") {
      setAdminRole("delivery");
      setAdminTab("orders");
      setView("admin");
      setShowAuthModal(false);
      setAuthFields({ email: "", password: "", name: "", phone: "" });
      showToast("Acceso concedido a Panel Delivery ✓");
      setAuthLoading(false);
      return;
    }

    // Check if password alone matches admin or delivery credentials as fallback
    if (authMode === "login" && (inputPass === "fellhonpm" || inputPass === "botifelldely")) {
      try {
        const res = await adminLoginMut.mutateAsync({ data: { password: inputPass } });
        if (res.ok) {
          const role = (res.role === "delivery" ? "delivery" : "full") as "full" | "delivery";
          setAdminRole(role);
          if (role === "delivery") setAdminTab("orders");
          setView("admin");
          setShowAuthModal(false);
          setAuthFields({ email: "", password: "", name: "", phone: "" });
          showToast(role === "delivery" ? "Acceso como Delivery ✓" : "Acceso como Administrador ✓");
          setAuthLoading(false);
          return;
        }
      } catch {}
    }

    try {
      const url = authMode === "login" ? "/api/customers/login" : "/api/customers/register";
      const body = authMode === "login"
        ? { email: authFields.email, password: authFields.password }
        : { email: authFields.email, password: authFields.password, name: authFields.name, phone: authFields.phone };
      const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      const data = await r.json();
      if (!r.ok) { setAuthError(data.message ?? "Error al procesar."); return; }
      try { localStorage.setItem("customer_token", data.token); } catch {}
      setCustomerToken(data.token);
      setCustomer(data.customer);
      setShowAuthModal(false);
      setAuthFields({ email: "", password: "", name: "", phone: "" });
      showToast(`¡Bienvenido, ${data.customer.name}!`);
    } catch { setAuthError("Error de conexión. Intenta de nuevo."); }
    finally { setAuthLoading(false); }
  };

  const handleCustomerLogout = () => {
    try { localStorage.removeItem("customer_token"); } catch {}
    setCustomerToken(null);
    setCustomer(null);
    setSavedCarts([]);
    setShowAccountPanel(false);
    showToast("Sesión cerrada");
  };

  const compressImage = (file: File, maxPx = 1800, quality = 0.9): Promise<Blob> =>
    new Promise((resolve) => {
      if (file.size < 150_000) { resolve(file); return; }
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
        canvas.getContext("2d")!.drawImage(img, 0, 0, width, height);
        const useWebP = canvas.toDataURL("image/webp").startsWith("data:image/webp");
        canvas.toBlob(
          (blob) => resolve(blob ?? file),
          useWebP ? "image/webp" : "image/jpeg",
          quality,
        );
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(file); };
      img.src = url;
    });

  const handleImageUpload = async (
    e: ChangeEvent<HTMLInputElement>,
    callback: (url: string) => void,
    options: { skipCompression?: boolean } = {},
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      showToast(options.skipCompression ? "Subiendo..." : "Comprimiendo y subiendo...");
      const compressed: Blob = options.skipCompression ? file : await compressImage(file);
      const ext = compressed.type === "image/webp" ? "webp" : "jpg";
      const uploadName = file.name.replace(/\.[^.]+$/, "") + "." + ext;
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const presignRes = await fetch(`${apiBase}/storage/uploads/request-url`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: uploadName,
          size: compressed.size,
          contentType: compressed.type,
        }),
      });
      if (!presignRes.ok) throw new Error("presign failed");
      const { uploadURL, objectPath } = await presignRes.json();
      const putRes = await fetch(uploadURL, {
        method: "PUT",
        headers: { "Content-Type": compressed.type },
        body: compressed,
      });
      if (!putRes.ok) throw new Error("upload failed");
      const servedUrl = `${apiBase}/storage${objectPath}`;
      callback(servedUrl);
      const savings = Math.round((1 - compressed.size / file.size) * 100);
      showToast(savings > 5 ? `Imagen subida (${savings}% más liviana)` : "Imagen subida");
    } catch (err) {
      console.error(err);
      showToast("Error al subir imagen");
    } finally {
      if (e.target) e.target.value = "";
    }
  };

  const resolveImageUrl = async (
    url: string,
    setFn: (resolved: string) => void,
  ) => {
    if (!url) return;
    const isDirect =
      url.includes("/storage/objects/") ||
      /\.(jpg|jpeg|png|webp|gif|avif|svg)(\?|$)/i.test(url) ||
      url.startsWith("data:");
    if (isDirect) return;
    try {
      showToast("Resolviendo imagen...");
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(
        `${apiBase}/resolve-image?url=${encodeURIComponent(url)}`,
      );
      if (!res.ok) throw new Error("no resolved");
      const { resolved } = await res.json();
      setFn(resolved);
      showToast("Imagen lista");
    } catch {
      showToast("No se pudo resolver el enlace — prueba el link directo a la imagen");
    }
  };

  const handleDeleteStorageImage = async (
    imageUrl: string,
    clearFn: () => void,
  ) => {
    if (!imageUrl.includes("/storage/objects/")) return;
    try {
      const marker = "/storage/objects/";
      const idx = imageUrl.indexOf(marker);
      const objectPath = imageUrl.slice(idx + "/storage".length);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      await fetch(`${apiBase}${objectPath}`, { method: "DELETE" });
      clearFn();
      showToast("Imagen eliminada");
    } catch {
      showToast("Error al eliminar imagen");
    }
  };

  const handleAddToCartClick = (product: Product) => {
    if (product.options && product.options.length > 0) {
      setOptionModalInfo({ product, selectedOption: product.options[0] });
    } else {
      processAddToCart(product, null);
    }
  };

  const processAddToCart = (product: Product, selectedOption: string | null) => {
    setCart((prev) => {
      const existing = prev.find(
        (item) => item.id === product.id && item.selectedOption === selectedOption,
      );
      if (existing) {
        return prev.map((item) =>
          item === existing ? { ...item, quantity: item.quantity + 1 } : item,
        );
      }
      return [
        ...prev,
        { ...product, quantity: 1, selectedOption, cartItemId: Date.now() },
      ];
    });
    setOptionModalInfo(null);
    showToast(`¡${product.name} añadido al carrito!`);
  };

  const addDeposit = (product: Product) => {
    const depositProduct: Product = {
      ...product,
      id: -product.id,
      name: `Depósito retornable - ${product.name}`,
      price: product.depositoAmount ?? 500,
      options: [],
      optionsTitle: "",
    };
    processAddToCart(depositProduct, null);
  };

  const updateQuantity = (cartItemId: number, delta: number) => {
    setCart((prev) =>
      prev
        .map((item) =>
          item.cartItemId === cartItemId
            ? { ...item, quantity: item.quantity + delta }
            : item,
        )
        .filter((item) => item.quantity > 0),
    );
  };

  const createOrderMut = useCreateOrder();

  const openCheckout = () => {
    if (cart.length === 0) return;
    if (customer) {
      setCheckoutForm(f => ({
        ...f,
        customerName: customer.name,
        phone: customer.phone ?? f.phone,
      }));
    }
    setCheckoutStep(1);
    setCheckoutOpen(true);
  };

  const submitCheckout = async (e: FormEvent) => {
    e.preventDefault();
    if (cart.length === 0) return;
    const { customerName, address, phone, notes, deliveryType } = checkoutForm;
    if (!customerName.trim() || !phone.trim()) {
      showToast("Completa nombre y teléfono");
      return;
    }
    if (deliveryType === "delivery" && !address.trim()) {
      showToast("Ingresa la dirección de entrega");
      return;
    }

    // Client-side validations before opening WhatsApp (avoid wasted messages)
    if (settings.contingencyMode) {
      showToast(settings.contingencyMessage || "La tienda está temporalmente en Modo Contingencia y no está recibiendo pedidos en este momento.");
      return;
    }
    if (!isStoreOpen) {
      showToast(`Estamos fuera de nuestro horario de atención. La recepción de pedidos se habilita a las ${settings.openTime ?? "11:00"}.`);
      return;
    }
    if (!deliveryMinimumMet) {
      showToast(`El monto mínimo para Delivery es $${deliveryMinimum.toLocaleString("es-CL")}. Agrega más productos para continuar.`);
      return;
    }

    const items = cart.map((item) => ({
      name: item.name,
      quantity: item.quantity,
      price: item.price,
      selectedOption: item.selectedOption ?? undefined,
    }));

    const tipoEntrega = deliveryType === "delivery" ? "🛵 Delivery" : "🏪 Retiro en tienda";
    let text = `¡Hola! Soy *${customerName}* y quiero hacer el siguiente pedido:\n\n`;
    cart.forEach((item) => {
      const optionText = item.selectedOption ? ` (${item.selectedOption})` : "";
      text += `👉 ${item.quantity}x ${item.name}${optionText} - $${(
        item.price * item.quantity
      ).toLocaleString("es-CL")}\n`;
    });
    text += `\n📦 *Tipo de entrega:* ${tipoEntrega}`;
    if (isDelivery && selectedLocation) {
      text += ` — ${selectedLocation.name}`;
    }
    text += `\n📍 *Dirección:* ${address}`;
    if (phone) text += `\n📞 *Teléfono:* ${phone}`;
    if (notes) text += `\n📝 *Notas:* ${notes}`;
    if (isDelivery) {
      text += `\n\n🛵 *Costo delivery:* $${DELIVERY_COST.toLocaleString("es-CL")}`;
    }
    if (discountSavings > 0) {
      text += `\n🏷️ *Descuento (${appliedDiscount?.code}):* -$${discountSavings.toLocaleString("es-CL")}`;
    }
    text += `\n\n💰 *Total: $${finalTotal.toLocaleString("es-CL")}*\n\n¿Me confirmas?`;

    // Abrir WhatsApp antes del await para que quede dentro del gesto del usuario.
    // Los navegadores móviles bloquean window.open() si se llama después de un await.
    const cleanWa = (settings.whatsapp ?? "").replace(/\D/g, "");
    try {
      window.open(
        `https://wa.me/${cleanWa}?text=${encodeURIComponent(text)}`,
        "_blank",
      );
    } catch {}

    setCart([]);
    setAppliedDiscount(null);
    setDiscountCode("");
    setCheckoutOpen(false);
    setIsCartOpen(false);
    setCheckoutForm({ customerName: "", address: "", phone: "", notes: "", deliveryType: "retiro" });
    showToast("Pedido enviado");

    try {
      await createOrderMut.mutateAsync({
        data: {
          customerName,
          address,
          phone,
          notes,
          items,
          subtotal: discountSavings > 0 ? cartTotal : undefined,
          discountCode: discountSavings > 0 ? (appliedDiscount?.code ?? undefined) : undefined,
          discountAmount: discountSavings > 0 ? discountSavings : undefined,
          total: finalTotal,
          deliveryType,
        },
      });
    } catch (err) {
      console.error(err);
    }
  };

  const handleAddCategory = () => {
    const name = newCategoryName.trim();
    if (!name || categories.includes(name)) return;
    setNewCategoryName("");
    showToast("Categoría añadida");
    createCategoryMut.mutate(
      { data: { name } },
      {
        onSuccess: (newCat) => {
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, categories: [...old.categories, newCat] } : old,
          );
          refreshMenu();
        },
        onError: refreshMenu,
      },
    );
  };
  const handleAddAisle = () => {
    const name = newAisleName.trim();
    if (!name || aisles.includes(name)) return;
    setNewAisleName("");
    showToast("Pasillo añadido");
    createAisleMut.mutate(
      { data: { name } },
      {
        onSuccess: (newAisle) => {
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, aisles: [...old.aisles, newAisle] } : old,
          );
          refreshMenu();
        },
        onError: refreshMenu,
      },
    );
  };
  const handleAddSubcat = () => {
    const name = newSubcatName.trim();
    if (!name || subcategories.includes(name)) return;
    setNewSubcatName("");
    showToast("Subcategoría añadida");
    createSubcategoryMut.mutate(
      { data: { name } },
      {
        onSuccess: (newSub) => {
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, subcategories: [...old.subcategories, newSub] } : old,
          );
          refreshMenu();
        },
        onError: refreshMenu,
      },
    );
  };

  const handleDragStart = (e: DragEvent<HTMLDivElement>, index: number) =>
    e.dataTransfer.setData("dragIndex", String(index));
  const handleDrop = (e: DragEvent<HTMLDivElement>, dropIndex: number) => {
    const dragIndex = parseInt(e.dataTransfer.getData("dragIndex"));
    const newOrder = [...categoriesData];
    const [dragged] = newOrder.splice(dragIndex, 1);
    newOrder.splice(dropIndex, 0, dragged);
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, categories: newOrder } : old,
    );
    reorderCategoriesMut.mutate(
      { data: { ids: newOrder.map((c) => c.id) } },
      { onSuccess: () => refreshMenu(), onError: refreshMenu },
    );
  };

  const moveAisle = (index: number, dir: -1 | 1) => {
    const target = index + dir;
    if (target < 0 || target >= aislesData.length) return;
    const newOrder = [...aislesData];
    [newOrder[index], newOrder[target]] = [newOrder[target], newOrder[index]];
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, aisles: newOrder } : old,
    );
    reorderAislesMut.mutate(
      { data: { ids: newOrder.map((a) => a.id) } },
      { onSuccess: () => refreshMenu(), onError: refreshMenu },
    );
  };

  const moveSubcat = (index: number, dir: -1 | 1) => {
    const target = index + dir;
    if (target < 0 || target >= subcategoriesData.length) return;
    const newOrder = [...subcategoriesData];
    [newOrder[index], newOrder[target]] = [newOrder[target], newOrder[index]];
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, subcategories: newOrder } : old,
    );
    reorderSubcategoriesMut.mutate(
      { data: { ids: newOrder.map((s) => s.id) } },
      { onSuccess: () => refreshMenu(), onError: refreshMenu },
    );
  };

  const startEditing = (product: Product) => {
    setEditingProduct(product.id);
    setProductWizardStep(1);
    setShowAddProductModal(false);
    setOpenProductFormSections({ main: true, media: false, extras: false });
    setProductFormSubTab("main");
    setFormState({
      name: product.name,
      price: String(product.price),
      image: product.image,
      category: product.category,
      aisle: product.aisle,
      subcategory: product.subcategory || "",
      optionsTitle: product.optionsTitle || "",
      optionsString: product.options ? product.options.join(", ") : "",
      oferta: product.oferta ?? false,
      depositoEnabled: product.depositoEnabled ?? false,
      depositoAmount: String(product.depositoAmount ?? 500),
      transferenciaEnabled: product.transferenciaEnabled ?? false,
      transferenciaAmount: String(product.transferenciaAmount ?? 0),
      contingencyEnabled: product.contingencyEnabled ?? false,
    });
  };

  const cancelEditing = () => {
    setEditingProduct(null);
    setShowAddProductModal(false);
    setProductWizardStep(1);
    setOpenProductFormSections({ main: true, media: false, extras: false });
    setProductFormSubTab("main");
    setFormState(DEFAULT_FORM);
  };

  const handleDownloadTemplate = () => {
    window.open(`${import.meta.env.BASE_URL}api/admin/download-template`, "_blank");
  };

  const handleDownloadProductsExcelWithImages = () => {
    showToast("Generando y descargando Excel con imágenes en máxima calidad...");
    const link = document.createElement("a");
    link.href = `${import.meta.env.BASE_URL}api/admin/export-excel-with-images`;
    link.download = `productos_fellas_con_imagenes_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      showToast("Procesando planilla Excel e importando...");
      const reader = new FileReader();
      reader.onload = async (event) => {
        const buffer = event.target?.result as ArrayBuffer;
        if (!buffer) return;
        let binary = "";
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const base64 = btoa(binary);

        const res = await fetch(`${import.meta.env.BASE_URL}api/admin/import-excel`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ base64 }),
        });
        const data = await res.json();
        if (data.ok) {
          const imgText = data.imagesExtracted > 0 ? ` y ${data.imagesExtracted} imágenes extraídas` : "";
          showToast(`¡Importados ${data.importedCount} productos con éxito${imgText}! Pasillos y categorías creados.`);
          refreshMenu();
        } else {
          showToast(data.error || "Error al importar Excel");
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error(err);
      showToast("Error al leer el archivo Excel");
    } finally {
      if (e.target) e.target.value = "";
    }
  };

  const handleExportBackupZip = () => {
    showToast("Generando copia de seguridad completa con imágenes (.ZIP)...");
    const link = document.createElement("a");
    link.href = `${import.meta.env.BASE_URL}api/admin/backup/export`;
    link.download = `fellas_backup_${new Date().toISOString().slice(0, 10)}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleExportBackupJson = () => {
    showToast("Descargando catálogo en JSON...");
    const link = document.createElement("a");
    link.href = `${import.meta.env.BASE_URL}api/admin/backup/export-json`;
    link.download = `fellas_catalogo_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleRestoreBackup = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      showToast(`Restaurando copia de seguridad desde ${file.name}...`);
      const reader = new FileReader();
      reader.onload = async (event) => {
        const buffer = event.target?.result as ArrayBuffer;
        if (!buffer) return;
        let binary = "";
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const base64 = btoa(binary);

        const res = await fetch(`${import.meta.env.BASE_URL}api/admin/backup/restore`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ base64, filename: file.name }),
        });
        const data = await res.json();
        if (data.ok) {
          showToast(data.message || "¡Copia de seguridad restaurada con éxito!");
          refreshMenu();
          loadQuickMediaImages();
        } else {
          showToast(data.error || "Error al restaurar copia de seguridad");
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error(err);
      showToast("Error al procesar el archivo de copia de seguridad");
    } finally {
      if (e.target) e.target.value = "";
    }
  };

  const handleDeleteAllProducts = async () => {
    if (!window.confirm("¿Estás 100% seguro de ELIMINAR TODOS LOS PRODUCTOS del catálogo? Esta acción no se puede deshacer.")) {
      return;
    }
    try {
      showToast("Eliminando todos los productos...");
      const res = await fetch(`${import.meta.env.BASE_URL}api/admin/products/all`, {
        method: "DELETE",
      });
      const data = await res.json();
      if (data.ok) {
        showToast("Todos los productos han sido eliminados.");
        refreshMenu();
      } else {
        showToast("Error al eliminar productos");
      }
    } catch {
      showToast("Error al eliminar productos");
    }
  };

  const handleLogin = async () => {
    try {
      const res = await adminLoginMut.mutateAsync({
        data: { password: passwordInput },
      });
      if (res.ok) {
        const role = (res.role === "delivery" ? "delivery" : "full") as "full" | "delivery";
        setAdminRole(role);
        if (role === "delivery") setAdminTab("orders");
        setView("admin");
        setPasswordInput("");
      } else {
        showToast("Contraseña incorrecta");
      }
    } catch {
      showToast("Contraseña incorrecta");
    }
  };

  const resolveImageForSave = async (url: string): Promise<string> => {
    if (!url) return url;
    const isDirect =
      url.includes("/storage/objects/") ||
      /\.(jpg|jpeg|png|webp|gif|avif|svg)(\?|$)/i.test(url) ||
      url.startsWith("data:");
    if (isDirect) return url;
    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/resolve-image?url=${encodeURIComponent(url)}`);
      if (!res.ok) return url;
      const { resolved } = await res.json();
      return resolved || url;
    } catch {
      return url;
    }
  };

  const saveProduct = async (e?: FormEvent | KeyboardEvent | React.SyntheticEvent) => {
    if (e && typeof e.preventDefault === "function") {
      e.preventDefault();
    }
    if (!formState.name || !formState.name.trim()) {
      showToast("Por favor ingresa el nombre del producto.");
      return;
    }
    if (!formState.category || !formState.aisle) {
      showToast("Por favor selecciona una categoría y un pasillo.");
      return;
    }
    const finalImage = await resolveImageForSave(formState.image);
    const parsedOptions = formState.optionsString
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s !== "");
    const productData = {
      name: formState.name.trim(),
      price: parseInt(formState.price) || 0,
      category: formState.category,
      aisle: formState.aisle,
      subcategory: formState.subcategory || "",
      optionsTitle: formState.optionsTitle || "",
      options: parsedOptions,
      image: finalImage || "",
      oferta: formState.oferta,
      depositoEnabled: formState.depositoEnabled,
      depositoAmount: parseInt(formState.depositoAmount) || 500,
      transferenciaEnabled: formState.transferenciaEnabled,
      transferenciaAmount: parseInt(formState.transferenciaAmount) || 0,
      contingencyEnabled: formState.contingencyEnabled,
    };
    if (editingProduct) {
      const id = editingProduct;
      queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
        old
          ? { ...old, products: old.products.map((p: any) => (p.id === id ? { ...p, ...productData } : p)) }
          : old,
      );
      showToast("Producto actualizado exitosamente");
      cancelEditing();
      updateProductMut.mutate({ id, data: productData }, { onSuccess: () => refreshMenu(), onError: refreshMenu });
    } else {
      showToast("Producto creado exitosamente");
      cancelEditing();
      createProductMut.mutate(
        { data: productData },
        {
          onSuccess: (newProduct) => {
            queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
              old ? { ...old, products: [...old.products, newProduct] } : old,
            );
            refreshMenu();
          },
          onError: refreshMenu,
        },
      );
    }
  };

  const deleteProductHandler = (id: number) => {
    const product = products.find((p) => p.id === id);
    if (!product) return;
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, products: old.products.filter((p: any) => p.id !== id) } : old,
    );
    setToast({
      show: true,
      message: "Producto eliminado",
      actionLabel: "Deshacer",
      onAction: async () => {
        await createProductMut.mutateAsync({
          data: {
            name: product.name,
            price: product.price,
            category: product.category,
            aisle: product.aisle,
            subcategory: product.subcategory || "",
            optionsTitle: product.optionsTitle || "",
            options: product.options || [],
            image: product.image,
          },
        });
        refreshMenu();
      },
    });
    deleteProductMut.mutate({ id }, { onSuccess: () => refreshMenu(), onError: refreshMenu });
  };

  const deleteCategoryHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, categories: old.categories.filter((c: any) => c.id !== id) } : old,
    );
    deleteCategoryMut.mutate({ id }, { onSuccess: () => refreshMenu(), onError: refreshMenu });
  };
  const deleteAisleHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, aisles: old.aisles.filter((a: any) => a.id !== id) } : old,
    );
    deleteAisleMut.mutate({ id }, { onSuccess: () => refreshMenu(), onError: refreshMenu });
  };
  const deleteSubcategoryHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, subcategories: old.subcategories.filter((s: any) => s.id !== id) } : old,
    );
    deleteSubcategoryMut.mutate({ id }, { onSuccess: () => refreshMenu(), onError: refreshMenu });
  };

  const saveSettings = () => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, settings: { ...old.settings, ...settingsDraft } } : old,
    );
    showToast("Ajustes guardados");
    updateSettingsMut.mutate(
      { data: settingsDraft },
      { onSuccess: () => refreshMenu(), onError: refreshMenu },
    );
  };

  if (!ageVerified) {
    return (
      <div className="min-h-screen bg-[#141414] flex items-center justify-center p-4">
        <div className="bg-[#1a1a1a] border border-[#ffd025]/20 rounded-3xl p-8 md:p-12 max-w-sm w-full flex flex-col items-center gap-7 shadow-2xl shadow-black/70">
          <div className="w-full flex justify-center min-h-[80px] items-center">
            {isLoading || !settings.logo ? (
              <div className="w-44 h-16 rounded-2xl bg-[#252525] animate-pulse" />
            ) : (
              <img
                src={settings.logo}
                alt="Logo"
                className="max-h-24 max-w-[200px] object-contain"
                onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
              />
            )}
          </div>
          <div className="w-full h-px bg-[#ffd025]/10" />
          <div className="text-center space-y-2">
            <p className="text-gray-300 text-base leading-snug">
              Debes ser mayor de 18 años para comprar en nuestra tienda
            </p>
            <p className="text-[#ffd025] text-xl font-black mt-2">
              ¿Eres mayor de edad?
            </p>
          </div>
          <div className="flex gap-4 w-full">
            <button
              onClick={handleAgeNo}
              className="flex-1 py-3 rounded-2xl border border-white/20 text-white font-bold text-lg hover:bg-white/5 transition-colors"
            >
              No
            </button>
            <button
              onClick={handleAgeYes}
              className="flex-1 py-3 rounded-2xl bg-[#ffd025] text-[#141414] font-black text-lg hover:brightness-110 transition-all shadow-lg shadow-[#ffd025]/20"
            >
              Sí
            </button>
          </div>
          <p className="text-gray-600 text-[11px] text-center leading-relaxed">
            Al ingresar confirmas que eres mayor de 18 años y aceptas la venta responsable de alcohol.
          </p>
        </div>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="min-h-screen bg-[#141414] flex items-center justify-center">
        <div className="w-10 h-10 rounded-full border-4 border-[#ffd025]/20 border-t-[#ffd025] animate-spin" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#141414] text-[#ffd025] font-sans selection:bg-[#ffd025] selection:text-[#141414]">
      {view === "client" && (
        <div className="fixed bottom-4 right-4 z-50 flex flex-col items-center gap-2.5">
          <a
            href="https://www.instagram.com/fellasmarketpm/"
            target="_blank"
            rel="noopener noreferrer"
            aria-label="Instagram"
            className="text-white p-3 rounded-full shadow-2xl hover:scale-110 active:scale-95 transition-transform border-2 border-[#141414]"
            style={{ background: "linear-gradient(135deg, #833ab4, #E1306C, #fd1d1d)" }}
          >
            <Instagram size={20} />
          </a>
          {showScrollTop && (
            <button
              onClick={scrollToTop}
              aria-label="Volver arriba"
              className="bg-[#ffd025] text-[#141414] p-3.5 rounded-full shadow-2xl shadow-[#ffd025]/40 hover:scale-110 active:scale-95 transition-transform border-2 border-[#141414]"
            >
              <ArrowUp size={20} strokeWidth={3} />
            </button>
          )}
        </div>
      )}
      {toast.show && (
        <div className="fixed top-24 right-4 z-50 bg-[#ffd025] text-[#141414] px-6 py-3 rounded-full font-bold shadow-lg shadow-[#ffd025]/20 flex items-center gap-3 animate-slide-in-right">
          <CheckCircle size={20} /> {toast.message}
          {toast.actionLabel && toast.onAction && (
            <button
              type="button"
              onClick={() => {
                toast.onAction?.();
                setToast({ show: false, message: "", actionLabel: "", onAction: null });
              }}
              className="ml-2 bg-[#141414] text-white px-3 py-1 rounded-full text-xs uppercase"
            >
              {toast.actionLabel}
            </button>
          )}
        </div>
      )}



      {/* Mobile Navigation Drawer */}
      {view === "client" && mobileNavDrawerOpen && (
        <div className="fixed inset-0 z-50 flex">
          <div
            className="fixed inset-0 bg-black/80 backdrop-blur-sm transition-opacity"
            onClick={() => setMobileNavDrawerOpen(false)}
          />
          <div className="relative w-4/5 max-w-xs bg-[#0d0d16] border-r border-white/10 h-full flex flex-col p-6 shadow-2xl z-10 overflow-y-auto animate-slide-in-right">
            <div className="flex items-center justify-between pb-4 border-b border-white/10">
              <div className="flex items-center gap-2">
                {settings.logo ? (
                  <img src={settings.logo} alt="Logo" className="h-7 w-auto object-contain" />
                ) : (
                  <span className="text-[#ffd025] font-black text-lg tracking-wider">FELLAS</span>
                )}
              </div>
              <button
                onClick={() => setMobileNavDrawerOpen(false)}
                className="p-1.5 text-gray-400 hover:text-white rounded-lg hover:bg-white/10"
              >
                <X size={20} />
              </button>
            </div>

            {/* Store status pill */}
            <div className="my-4 p-3 rounded-2xl bg-white/5 border border-white/10 flex items-center justify-between text-xs">
              <span className="text-gray-400 font-medium">Horario</span>
              <span className={`font-bold flex items-center gap-1.5 ${isStoreOpen ? "text-emerald-400" : "text-amber-400"}`}>
                <span className={`w-2 h-2 rounded-full ${isStoreOpen ? "bg-emerald-400 animate-pulse" : "bg-amber-400"}`} />
                {isStoreOpen ? "Abierto" : `Abre ${settings.openTime ?? "11:00"}`}
              </span>
            </div>

            {/* Direct contact actions */}
            <div className="space-y-2 mb-5">
              {settings.whatsapp && (
                <a
                  href={`https://wa.me/${settings.whatsapp.replace(/\D/g, "")}`}
                  target="_blank"
                  rel="noreferrer"
                  className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-300 text-xs font-bold hover:bg-emerald-500/20 transition-colors"
                >
                  <Phone size={15} /> WhatsApp Tienda
                </a>
              )}
              {settings.contactPhone && (
                <a
                  href={`tel:${settings.contactPhone}`}
                  className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl bg-white/5 border border-white/10 text-gray-300 text-xs font-bold hover:text-white transition-colors"
                >
                  <PhoneCall size={15} className="text-[#ffd025]" /> {settings.contactPhone}
                </a>
              )}
            </div>

            {/* Pasillos & Categorías list */}
            <div className="flex-1 space-y-4">
              <p className="text-[11px] font-black uppercase text-[#ffd025] tracking-widest">Navegación de Pasillos</p>
              
              {/* Pestañas Especiales Oportunidades & Packs en Móvil */}
              <div className="grid grid-cols-2 gap-2 mb-2">
                <button
                  onClick={() => {
                    setSearchQuery("");
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("oportunidades");
                    setMobileNavDrawerOpen(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`px-3 py-2.5 rounded-xl text-xs font-black uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 border ${
                    showDedicatedProductsPage && dedicatedViewMode === "oportunidades"
                      ? "bg-amber-400 text-black border-amber-400 shadow"
                      : "bg-amber-500/10 text-amber-300 border-amber-500/30 hover:bg-amber-500/20"
                  }`}
                >
                  <span>OPORTUNIDADES</span>
                  <span>⏰</span>
                </button>
                <button
                  onClick={() => {
                    setSearchQuery("");
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("packs");
                    setMobileNavDrawerOpen(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`px-3 py-2.5 rounded-xl text-xs font-black uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 border ${
                    showDedicatedProductsPage && dedicatedViewMode === "packs"
                      ? "bg-amber-400 text-black border-amber-400 shadow"
                      : "bg-purple-500/10 text-purple-300 border-purple-500/30 hover:bg-purple-500/20"
                  }`}
                >
                  <span>PACKS</span>
                  <span>🎁</span>
                </button>
              </div>

              <div className="space-y-1">
                <button
                  onClick={() => {
                    setActiveCategory("");
                    setActiveAisle("");
                    setShowDedicatedProductsPage(false);
                    setDedicatedViewMode("catalog");
                    setMobileNavDrawerOpen(false);
                  }}
                  className={`w-full text-left px-3.5 py-2 rounded-xl text-xs font-bold transition-colors flex items-center justify-between ${
                    !showDedicatedProductsPage && !activeCategory && !activeAisle ? "bg-[#ffd025] text-black" : "text-gray-300 hover:bg-white/5"
                  }`}
                >
                  <span>Todos los productos</span>
                  <span className="text-[10px] opacity-75">{baseProducts.length}</span>
                </button>
                {allMenuSections.map((sec) => {
                  const isActive = sec.type === "category" ? activeCategory === sec.name : activeAisle === sec.name;
                  const count = baseProducts.filter((p) => (sec.type === "category" ? p.category === sec.name : p.aisle === sec.name)).length;
                  return (
                    <button
                      key={sec.name}
                      onClick={() => {
                        setNavQuickFilter("");
                        if (sec.type === "category") {
                          setActiveAisle("");
                          setActiveCategory(sec.name);
                        } else {
                          setActiveCategory("");
                          setActiveAisle(sec.name);
                        }
                        setShowDedicatedProductsPage(true);
                        setDedicatedViewMode("pasillo");
                        setMobileNavDrawerOpen(false);
                        window.scrollTo({ top: 0, behavior: "smooth" });
                      }}
                      className={`w-full text-left px-3.5 py-2 rounded-xl text-xs font-bold transition-colors flex items-center justify-between ${
                        isActive ? "bg-[#ffd025] text-black" : "text-gray-300 hover:bg-white/5"
                      }`}
                    >
                      <span className="truncate">{sec.name}</span>
                      <span className="text-[10px] px-1.5 py-0.5 rounded-md bg-white/10 text-gray-400 font-mono">{count}</span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Bottom account / admin actions */}
            <div className="pt-4 border-t border-white/10 space-y-2 mt-4">
              <button
                onClick={() => {
                  setMobileNavDrawerOpen(false);
                  if (customer) setShowAccountPanel(true);
                  else { setShowAuthModal(true); setAuthMode("login"); }
                }}
                className="w-full flex items-center gap-2 px-3.5 py-2.5 rounded-xl bg-white/5 hover:bg-white/10 text-xs font-bold text-gray-200 transition-colors"
              >
                <User size={15} className="text-[#ffd025]" />
                <span>{customer ? customer.name : "Iniciar Sesión / Registro"}</span>
              </button>
              <button
                onClick={() => {
                  setMobileNavDrawerOpen(false);
                  setView("admin-login");
                }}
                className="w-full flex items-center gap-2 px-3.5 py-2.5 rounded-xl text-xs font-bold text-gray-400 hover:text-white transition-colors"
              >
                <SettingsIcon size={15} /> Panel de Administración
              </button>
            </div>
          </div>
        </div>
      )}

      {view === "client" && (
        <header className="sticky top-0 z-50 bg-black text-white shadow-2xl select-none">
          {/* BARRA SUPERIOR (1/6): Degradado naranja-coral ultra delgado con tipografía equilibrada */}
          <div
            onClick={() => setShowComunasModal(true)}
            className="w-full text-white py-1.5 px-2 sm:px-4 text-center cursor-pointer select-none transition-opacity hover:opacity-95 flex items-center justify-center min-h-[24px] sm:min-h-[28px]"
            style={{
              background: "linear-gradient(90deg, #f7a627 0%, #fa7a34 50%, #f44369 100%)",
            }}
          >
            <span className="font-bold uppercase tracking-wider text-[7.5px] min-[360px]:text-[8.5px] min-[410px]:text-[9.5px] sm:text-[11px] leading-tight text-white text-center break-words text-balance max-w-full">
              {settings.topAnnouncementText?.trim() || "PIDE ANTES DE LAS 8:00 AM Y RECIBE EL MISMO DÍA (VER COMUNAS)"}
            </span>
          </div>

          {/* BARRA PRINCIPAL (5/6): División continua con líneas separadoras/divisorias verticales */}
          <div className="w-full bg-black border-b border-white/20">
            <div className="w-full flex items-stretch h-12 sm:h-14 px-2 sm:px-6 md:px-8 lg:px-10">
              
              {/* 1. SECCIÓN IZQUIERDA: LOGO + CATEGORÍAS (ÍCONO EN MÓVIL) + OPORTUNIDADES + PACKS */}
              <div className="flex items-center gap-2 sm:gap-5 pr-2 sm:pr-6 shrink-0">
                {/* LOGO LA NEGRA */}
                <button
                  className="focus:outline-none flex items-center group mr-0.5 sm:mr-2"
                  onClick={() => {
                    setSearchQuery("");
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  title="Ir al inicio"
                >
                  {settings.logo ? (
                    <img
                      src={settings.logo}
                      alt="LA NEGRA"
                      className="h-6 sm:h-8 md:h-9 object-contain max-w-[80px] sm:max-w-[140px] md:max-w-[180px] brightness-110"
                      onError={(e) => {
                        (e.target as HTMLImageElement).style.display = "none";
                      }}
                    />
                  ) : null}
                  {!settings.logo && (
                    <span className="font-serif font-black text-base sm:text-2xl md:text-[26px] tracking-tight text-white uppercase whitespace-nowrap">
                      LA NEGRA
                    </span>
                  )}
                </button>

                {/* CATEGORÍAS (Solo ícono en PC y móvil) */}
                <div className="relative h-full flex items-center">
                  <button
                    onClick={() => {
                      setShowAuthBar(false);
                      setShowLocationBar(false);
                      setIsCartOpen(false);
                      setShowAisleMenu((prev) => !prev);
                    }}
                    className={`flex items-center justify-center p-1.5 sm:p-2 rounded transition-colors ${
                      showAisleMenu || activeCategory || activeAisle
                        ? "text-[#ffd025]"
                        : "text-white hover:text-gray-300"
                    }`}
                    title="Categorías y Pasillos"
                    aria-expanded={showAisleMenu}
                  >
                    <LayoutGrid size={19} className="shrink-0" />
                  </button>
                </div>

                {/* OPORTUNIDADES ⏰ (Visible en pantallas grandes) */}
                <button
                  onClick={() => {
                    setSearchQuery("");
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("oportunidades");
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`hidden lg:flex items-center gap-1.5 text-xs sm:text-sm font-black tracking-wider uppercase transition-colors shrink-0 ${
                    showDedicatedProductsPage && dedicatedViewMode === "oportunidades"
                      ? "text-amber-400"
                      : "text-white hover:text-gray-300"
                  }`}
                  title="Ver oportunidades y ofertas"
                >
                  <span>OPORTUNIDADES</span>
                  <span className="text-sm">⏰</span>
                </button>

                {/* PACKS (Visible en pantallas grandes) */}
                <button
                  onClick={() => {
                    setSearchQuery("");
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("packs");
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`hidden lg:block text-xs sm:text-sm font-black tracking-wider uppercase transition-colors shrink-0 ${
                    showDedicatedProductsPage && dedicatedViewMode === "packs"
                      ? "text-amber-400"
                      : "text-white hover:text-gray-300"
                  }`}
                  title="Ver packs y promociones"
                >
                  PACKS
                </button>
              </div>

              {/* 2. SECCIÓN CENTRAL: DIVISIÓN DE BÚSQUEDA CON LÍNEAS SEPARADORAS A AMBOS LADOS */}
              <div className="flex-1 flex items-center h-full px-2 sm:px-6 min-w-0 border-l border-r border-white/20">
                <div className="w-full h-full flex items-center relative">
                  <Search
                    size={16}
                    className="text-white shrink-0 mr-1.5 sm:mr-3 pointer-events-none opacity-90"
                    strokeWidth={2.2}
                  />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="¿QUÉ BUSCAS?"
                    className="w-full h-full bg-transparent text-[11px] sm:text-sm font-bold uppercase tracking-wider text-white placeholder-gray-400 focus:outline-none truncate"
                  />
                  {searchQuery && (
                    <button
                      onClick={() => setSearchQuery("")}
                      className="text-gray-400 hover:text-white p-1 shrink-0"
                      aria-label="Limpiar búsqueda"
                    >
                      <X size={15} />
                    </button>
                  )}
                </div>
              </div>

              {/* 3. SECCIÓN DERECHA: BOTONES EN FORMATO SOLO ÍCONO CON LÍNEAS SEPARADORAS */}
              <div className="flex items-stretch shrink-0">
                {/* INICIA SESIÓN / MI CUENTA (Solo ícono) */}
                <button
                  onClick={() => {
                    setShowAisleMenu(false);
                    setShowLocationBar(false);
                    setIsCartOpen(false);
                    setShowAuthBar((prev) => !prev);
                  }}
                  className={`flex items-center justify-center px-3 sm:px-4 h-full border-r border-white/20 transition-colors shrink-0 ${
                    showAuthBar ? "text-[#ffd025]" : "text-white hover:text-gray-300"
                  }`}
                  title={customer ? `Sesión activa: ${customer.name}` : "Inicia Sesión / Mi Cuenta"}
                  aria-label="Cuenta de usuario"
                >
                  <User size={19} className="shrink-0" strokeWidth={1.8} />
                </button>

                {/* ELECCIÓN DE UBICACIÓN / COMUNAS (Solo ícono - Solo visible en versión PC) */}
                <button
                  onClick={() => {
                    setShowAisleMenu(false);
                    setShowAuthBar(false);
                    setIsCartOpen(false);
                    setShowLocationBar((prev) => !prev);
                  }}
                  className={`hidden md:flex items-center justify-center px-3 sm:px-4 h-full text-white border-r border-white/20 transition-colors shrink-0 ${
                    showLocationBar ? "text-[#ffd025]" : "hover:text-[#ffd025]"
                  }`}
                  title={selectedComuna ? `Comuna seleccionada: ${selectedComuna}` : "Seleccionar comuna de entrega"}
                  aria-label="Seleccionar comuna de entrega"
                >
                  <MapPin
                    size={19}
                    className={showLocationBar || selectedComuna ? "text-[#ffd025] shrink-0" : "text-white shrink-0"}
                    strokeWidth={1.8}
                  />
                </button>

                {/* CARRITO */}
                <button
                  onClick={() => {
                    setShowAisleMenu(false);
                    setShowAuthBar(false);
                    setShowLocationBar(false);
                    setIsCartOpen((prev) => !prev);
                  }}
                  className={`flex items-center justify-center gap-1 px-2 sm:px-4 h-full transition-all shrink-0 ${
                    isCartOpen ? "text-[#ffd025]" : "text-white hover:text-gray-300"
                  }`}
                  aria-label="Ver carrito"
                  title="Ver carrito de compras"
                >
                  <ShoppingCart size={20} className="shrink-0" strokeWidth={1.8} />
                  <span className={`w-4.5 h-4.5 sm:w-5 sm:h-5 rounded-full border text-[10px] sm:text-[11px] font-bold flex items-center justify-center shrink-0 ${
                    isCartOpen ? "border-[#ffd025] text-[#ffd025]" : "border-white text-white"
                  }`}>
                    {cartItemCount}
                  </span>
                </button>
              </div>

            </div>
          </div>

          {/* 4. DESPLIEGUE HORIZONTAL DE PASILLOS Y CATEGORÍAS (TEXTO SUELTO ULTRA COMPACTO) */}
          {showAisleMenu && (
            <div className="w-full bg-[#0a0a0f] border-b border-white/10 shadow-md py-1 px-2.5 sm:px-4 animate-fade-in transition-all">
              <div className="max-w-[1500px] mx-auto flex items-center gap-3.5 sm:gap-5 overflow-x-auto no-scrollbar scrollbar-none">
                <button
                  onClick={() => {
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(false);
                    setDedicatedViewMode("catalog");
                    setShowAisleMenu(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`text-[10.5px] sm:text-xs uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 py-0.5 ${
                    !activeCategory && !activeAisle && !navQuickFilter && (!showDedicatedProductsPage || dedicatedViewMode === "catalog")
                      ? "text-[#ffd025] font-black"
                      : "text-gray-400 hover:text-white font-semibold"
                  }`}
                >
                  Todo el catálogo
                </button>

                <button
                  onClick={() => {
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("oportunidades");
                    setShowAisleMenu(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`text-[10.5px] sm:text-xs uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 py-0.5 ${
                    showDedicatedProductsPage && dedicatedViewMode === "oportunidades"
                      ? "text-amber-400 font-black"
                      : "text-amber-300/80 hover:text-amber-300 font-semibold"
                  }`}
                >
                  ⏰ Oportunidades
                </button>

                <button
                  onClick={() => {
                    setActiveCategory("");
                    setActiveAisle("");
                    setNavQuickFilter("");
                    setShowDedicatedProductsPage(true);
                    setDedicatedViewMode("packs");
                    setShowAisleMenu(false);
                    window.scrollTo({ top: 0, behavior: "smooth" });
                  }}
                  className={`text-[10.5px] sm:text-xs uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 py-0.5 ${
                    showDedicatedProductsPage && dedicatedViewMode === "packs"
                      ? "text-amber-400 font-black"
                      : "text-amber-300/80 hover:text-amber-300 font-semibold"
                  }`}
                >
                  🎁 Packs
                </button>

                {allMenuSections.map((item) => {
                  const isActive =
                    item.type === "category"
                      ? activeCategory === item.name
                      : activeAisle === item.name;

                  return (
                    <button
                      key={item.name}
                      onClick={() => {
                        setNavQuickFilter("");
                        if (item.type === "category") {
                          setActiveAisle("");
                          setActiveCategory(item.name);
                        } else {
                          setActiveCategory("");
                          setActiveAisle(item.name);
                        }
                        setShowDedicatedProductsPage(true);
                        setDedicatedViewMode("pasillo");
                        setShowAisleMenu(false);
                        window.scrollTo({ top: 0, behavior: "smooth" });
                      }}
                      className={`text-[10.5px] sm:text-xs uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 py-0.5 ${
                        isActive
                          ? "text-[#ffd025] font-black"
                          : "text-gray-400 hover:text-white font-semibold"
                      }`}
                    >
                      {item.name}
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* 5. DESPLIEGUE HORIZONTAL DE INICIO DE SESIÓN Y CUENTA (CENTRADO EN MÓVIL Y ESCRITORIO) */}
          {showAuthBar && (
            <div className="w-full bg-[#0a0a0f] border-b border-white/10 shadow-md py-1 px-2 sm:px-4 animate-fade-in transition-all">
              <div className="max-w-[1500px] mx-auto flex items-center justify-center gap-2 sm:gap-3 overflow-x-auto no-scrollbar scrollbar-none">
                {!customer ? (
                  <form
                    onSubmit={(e) => {
                      e.preventDefault();
                      setAuthMode("login");
                      handleCustomerAuth(e);
                    }}
                    className="flex items-center justify-center gap-1.5 sm:gap-3 w-full overflow-x-auto no-scrollbar scrollbar-none py-0.5"
                  >
                    <input
                      type="text"
                      required
                      value={authFields.email}
                      onChange={(e) => {
                        setAuthFields((f) => ({ ...f, email: e.target.value }));
                        if (authError) setAuthError("");
                      }}
                      placeholder="Usuario / correo"
                      className="bg-white/10 border border-white/20 text-white rounded px-2 sm:px-2.5 py-0.5 text-[10.5px] sm:text-xs placeholder-gray-400 focus:outline-none focus:border-[#ffd025] w-24 xs:w-28 sm:w-44"
                    />

                    <input
                      type="password"
                      required
                      value={authFields.password}
                      onChange={(e) => {
                        setAuthFields((f) => ({ ...f, password: e.target.value }));
                        if (authError) setAuthError("");
                      }}
                      placeholder="Clave"
                      className="bg-white/10 border border-white/20 text-white rounded px-2 sm:px-2.5 py-0.5 text-[10.5px] sm:text-xs placeholder-gray-400 focus:outline-none focus:border-[#ffd025] w-20 xs:w-24 sm:w-32"
                    />

                    <button
                      type="submit"
                      disabled={authLoading}
                      className="text-[10.5px] sm:text-xs font-black uppercase text-[#ffd025] hover:brightness-125 px-1.5 sm:px-2 py-0.5 whitespace-nowrap shrink-0 transition-colors disabled:opacity-50"
                    >
                      {authLoading ? "..." : "Entrar"}
                    </button>

                    <span className="text-gray-600 text-[10px] select-none">|</span>

                    <button
                      type="button"
                      onClick={() => {
                        setShowAuthModal(true);
                        setAuthMode("register");
                        setShowAuthBar(false);
                      }}
                      className="text-gray-400 hover:text-[#ffd025] p-1 shrink-0 transition-colors"
                      title="Registrarse / Crear cuenta"
                      aria-label="Registrarse"
                    >
                      <UserPlus size={15} />
                    </button>

                    {authError && (
                      <span className="text-rose-400 text-[10px] font-bold whitespace-nowrap pl-1">
                        {authError}
                      </span>
                    )}
                  </form>
                ) : (
                  <div className="flex items-center justify-center gap-2.5 sm:gap-3 w-full overflow-x-auto no-scrollbar scrollbar-none py-0.5">
                    <span className="text-[10.5px] sm:text-xs uppercase font-black text-[#ffd025] whitespace-nowrap">
                      Hola, {customer.name}
                    </span>

                    <span className="text-gray-500 text-[10px] select-none">|</span>

                    <button
                      onClick={() => {
                        setShowAccountPanel(true);
                        setShowAuthBar(false);
                      }}
                      className="text-[10.5px] sm:text-xs font-semibold uppercase text-gray-300 hover:text-white whitespace-nowrap"
                    >
                      Carritos ({savedCarts.length})
                    </button>

                    <span className="text-gray-500 text-[10px] select-none">|</span>

                    <button
                      onClick={() => {
                        handleCustomerLogout();
                        setShowAuthBar(false);
                      }}
                      className="text-[10.5px] sm:text-xs font-bold uppercase text-red-400 hover:text-red-300 whitespace-nowrap"
                    >
                      Cerrar sesión
                    </button>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 5.1 DESPLIEGUE HORIZONTAL DE UBICACIÓN Y COMUNAS (SOLO EN PC, TEXTO SUELTO ULTRA COMPACTO Y CENTRADO) */}
          {showLocationBar && (
            <div className="hidden md:block w-full bg-[#0a0a0f] border-b border-white/10 shadow-md py-1 px-2.5 sm:px-4 animate-fade-in transition-all">
              <div className="max-w-[1500px] mx-auto flex items-center justify-center gap-3.5 sm:gap-5 overflow-x-auto no-scrollbar scrollbar-none">
                {[
                  "Santiago Centro",
                  "Providencia",
                  "Las Condes",
                  "Ñuñoa",
                  "Vitacura",
                  "La Florida",
                  "Maipú",
                  "San Miguel",
                  "Macul",
                  "Peñalolén",
                  "La Reina",
                  "Huechuraba",
                ].map((comuna) => {
                  const isSelected = selectedComuna === comuna;
                  return (
                    <button
                      key={comuna}
                      onClick={() => {
                        setSelectedComuna(comuna);
                        setCheckoutForm((prev) => ({
                          ...prev,
                          address: prev.address ? `${prev.address}, ${comuna}` : comuna,
                        }));
                        setShowLocationBar(false);
                        setToast({
                          show: true,
                          message: `📍 Ubicación de entrega: ${comuna}`,
                          actionLabel: "",
                          onAction: null,
                        });
                      }}
                      className={`text-[10.5px] sm:text-xs uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 py-0.5 ${
                        isSelected
                          ? "text-[#ffd025] font-black"
                          : "text-gray-400 hover:text-white font-semibold"
                      }`}
                    >
                      {comuna}
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* 6. DESPLIEGUE VERTICAL PLEGABLE DEL CARRITO JUSTO ABAJO DEL ENCABEZADO */}
          {isCartOpen && (
            <div className="w-full bg-[#0c0c12] border-b border-[#ffd025]/30 shadow-2xl animate-fade-in transition-all">
              <div className="max-w-[1500px] mx-auto px-3 sm:px-6 py-3.5">
                <div className="flex items-center justify-between pb-2.5 border-b border-white/10 mb-3">
                  <div className="flex items-center gap-2">
                    <ShoppingCart size={16} className="text-[#ffd025]" />
                    <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                      Tu Carrito ({cartItemCount} {cartItemCount === 1 ? "producto" : "productos"})
                    </span>
                  </div>
                  <button
                    onClick={() => setIsCartOpen(false)}
                    className="text-gray-400 hover:text-white p-1 transition-colors"
                    aria-label="Cerrar carrito"
                  >
                    <X size={16} />
                  </button>
                </div>

                {cart.length === 0 ? (
                  <div className="py-6 text-center text-gray-400">
                    <p className="text-xs font-bold uppercase tracking-wider text-gray-300">Tu carrito está vacío</p>
                    <p className="text-[11px] text-gray-500 mt-1">Explora los pasillos y añade tus productos</p>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {/* Lista vertical de productos en el carrito */}
                    <div className="max-h-64 sm:max-h-80 overflow-y-auto pr-1 space-y-2 no-scrollbar">
                      {cart.map((item) => (
                        <div
                          key={item.cartItemId}
                          className="flex items-center justify-between gap-3 bg-white/5 border border-white/10 rounded-xl p-2.5 transition-colors hover:bg-white/[0.08]"
                        >
                          <div className="flex items-center gap-3 min-w-0">
                            {item.image ? (
                              <img
                                src={item.image}
                                alt={item.name}
                                className="w-11 h-11 rounded-lg object-cover bg-black/40 shrink-0"
                                onError={(e) => {
                                  (e.currentTarget as HTMLImageElement).style.display = "none";
                                }}
                              />
                            ) : (
                              <div className="w-11 h-11 rounded-lg bg-white/10 flex items-center justify-center shrink-0">
                                <Package size={16} className="text-gray-400" />
                              </div>
                            )}
                            <div className="min-w-0">
                              <h4 className="font-bold text-white text-xs truncate leading-tight">
                                {item.name}
                              </h4>
                              {item.selectedOption && (
                                <p className="text-[10px] text-gray-400 truncate mt-0.5">
                                  {item.selectedOption}
                                </p>
                              )}
                              <p className="text-[#ffd025] font-black text-xs mt-0.5">
                                ${(Number(item.price || 0) * item.quantity).toLocaleString("es-CL")}
                              </p>
                            </div>
                          </div>

                          <div className="flex items-center gap-1 bg-black/60 rounded-lg p-1 border border-white/10 shrink-0">
                            <button
                              onClick={() => updateQuantity(item.cartItemId, -1)}
                              className="w-6 h-6 flex items-center justify-center rounded text-[#ffd025] hover:bg-[#ffd025]/20 font-bold text-xs transition-colors"
                            >
                              -
                            </button>
                            <span className="font-bold text-xs text-white w-5 text-center">
                              {item.quantity}
                            </span>
                            <button
                              onClick={() => updateQuantity(item.cartItemId, 1)}
                              className="w-6 h-6 flex items-center justify-center rounded text-[#ffd025] hover:bg-[#ffd025]/20 font-bold text-xs transition-colors"
                            >
                              +
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>

                    {/* Resumen y Botón de Enviar Pedido */}
                    <div className="pt-2.5 border-t border-white/10 flex flex-col sm:flex-row items-center justify-between gap-3">
                      <div className="flex items-baseline gap-2 w-full sm:w-auto justify-between sm:justify-start">
                        <span className="text-[11px] font-bold uppercase text-gray-400">Total:</span>
                        <span className="text-base sm:text-lg font-black text-[#ffd025]">
                          ${cartTotal.toLocaleString("es-CL")}
                        </span>
                      </div>

                      {!isStoreOpen && (
                        <p className="text-[10px] text-amber-400 font-bold text-center">
                          🕐 Pedidos habilitados desde las {settings.openTime ?? "11:00"}
                        </p>
                      )}

                      <button
                        onClick={() => {
                          setIsCartOpen(false);
                          openCheckout();
                        }}
                        disabled={!isStoreOpen}
                        className="w-full sm:w-auto px-6 py-2.5 bg-[#ffd025] text-black rounded-xl font-black text-xs uppercase tracking-wider hover:bg-[#e5b81a] transition-all flex items-center justify-center gap-2 shadow-lg disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        <Phone size={14} /> Enviar Pedido
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 3ra DIVISIÓN DEL ENCABEZADO: Aviso de Horario de Atención (ESTRICTAMENTE 1 LÍNEA) */}
          {!isStoreOpen && (
            <div className="w-full bg-[#1a0f00] border-t border-b border-amber-600/30 py-1 sm:py-1.5 px-2 sm:px-4 shadow-md overflow-hidden select-none">
              <div className="w-full max-w-[1500px] mx-auto flex items-center justify-center gap-1 sm:gap-2 whitespace-nowrap overflow-hidden">
                <span className="text-amber-400 text-[10px] sm:text-xs leading-none shrink-0 animate-pulse">🕐</span>
                <p className="text-[7.5px] min-[330px]:text-[8.5px] min-[380px]:text-[9.5px] min-[420px]:text-[10px] sm:text-[11.5px] md:text-xs text-amber-300 font-bold leading-tight m-0 text-center whitespace-nowrap truncate">
                  <span>Estamos fuera de nuestro horario de atención. </span>
                  <span className="text-gray-300 font-normal">Revisa el catálogo y arma tu carrito (pedidos desde las </span>
                  <span className="text-[#ffd025] font-black">{settings.openTime ?? "11:00"}</span>
                  <span className="text-gray-300 font-normal">)</span>
                </p>
              </div>
            </div>
          )}
        </header>
      )}

      {view === "client" && (
        <div className="w-full">
        <main className="relative">
          {optionModalInfo && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
              <div
                className="absolute inset-0 bg-black/80 backdrop-blur-sm"
                onClick={() => setOptionModalInfo(null)}
              ></div>
              <div className="relative bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl p-6 w-full max-w-sm shadow-2xl animate-fade-in">
                <button
                  onClick={() => setOptionModalInfo(null)}
                  className="absolute top-4 right-4 text-gray-400 hover:text-white"
                >
                  <X size={24} />
                </button>
                <h3 className="text-xl font-black text-white mb-1">
                  {optionModalInfo.product.name}
                </h3>
                <p className="text-[#ffd025] font-bold mb-6">
                  ${Number(optionModalInfo.product.price || 0).toLocaleString("es-CL")}
                </p>
                <div className="mb-6">
                  <label className="block text-sm font-bold text-gray-400 uppercase mb-3">
                    {optionModalInfo.product.optionsTitle || "Elige una opción:"}
                  </label>
                  <div className="space-y-2">
                    {optionModalInfo.product.options.map((opt) => (
                      <label
                        key={opt}
                        className={`flex items-center gap-3 p-3 rounded-xl border cursor-pointer transition-colors ${
                          optionModalInfo.selectedOption === opt
                            ? "border-[#ffd025] bg-[#ffd025]/10"
                            : "border-gray-700 hover:border-gray-500"
                        }`}
                      >
                        <input
                          type="radio"
                          name="product_option"
                          value={opt}
                          checked={optionModalInfo.selectedOption === opt}
                          onChange={() =>
                            setOptionModalInfo({
                              ...optionModalInfo,
                              selectedOption: opt,
                            })
                          }
                          className="accent-[#ffd025] w-4 h-4"
                        />
                        <span className="text-white text-sm font-medium">{opt}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <button
                  onClick={() =>
                    processAddToCart(
                      optionModalInfo.product,
                      optionModalInfo.selectedOption,
                    )
                  }
                  className="w-full py-3 bg-[#ffd025] text-[#141414] rounded-xl font-bold uppercase hover:bg-[#e5b81a] transition-all shadow-lg"
                >
                  Confirmar y Añadir
                </button>
              </div>
            </div>
          )}

          {/* RENDERIZADO CONDICIONAL: 1. BÚSQUEDA EXCLUSIVA (PRIORIDAD MÁXIMA) vs 2. PESTAÑA DEDICADA vs 3. HOME */}
          {searchQuery.trim().length > 0 ? (
            /* VISTA EXCLUSIVA DE RESULTADOS DE BÚSQUEDA (Sin Colecciones, Sin Recomendados, Sin Banners) */
            <section className="w-full px-3 sm:px-6 md:px-8 mt-4 sm:mt-6 mb-12 animate-fade-in">
              {/* Header de Búsqueda */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 mb-6 border-b border-white/10">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <Search size={18} className="text-[#ffd025]" />
                    <h2 className="text-base sm:text-lg font-black uppercase text-white tracking-wider">
                      Resultados de Búsqueda
                    </h2>
                    <span className="text-[10px] px-2.5 py-0.5 rounded-full bg-[#ffd025]/20 text-[#ffd025] font-bold font-mono border border-[#ffd025]/30">
                      {searchResults.length} {searchResults.length === 1 ? "producto encontrado" : "productos encontrados"}
                    </span>
                  </div>
                  <p className="text-xs text-gray-400 mt-1">
                    Mostrando resultados que coinciden con: <span className="text-white font-bold">"{searchQuery}"</span>
                  </p>
                </div>

                <button
                  type="button"
                  onClick={() => setSearchQuery("")}
                  className="self-start sm:self-auto text-gray-400 hover:text-white text-xs font-bold uppercase tracking-wider transition flex items-center gap-1.5 border-b border-white/20 hover:border-[#ffd025] pb-0.5 cursor-pointer"
                >
                  <X size={14} />
                  <span>Limpiar búsqueda</span>
                </button>
              </div>

              {/* Grilla Directa de Productos Encontrados */}
              {searchResults.length === 0 ? (
                <div className="text-center py-16 px-4 border border-dashed border-white/10 rounded-3xl bg-white/[0.02]">
                  <div className="w-16 h-16 rounded-2xl bg-white/5 border border-white/10 flex items-center justify-center mx-auto mb-3 text-gray-500">
                    <Search size={28} />
                  </div>
                  <h3 className="text-base sm:text-lg font-bold text-white mb-1">
                    No se encontraron productos para "{searchQuery}"
                  </h3>
                  <p className="text-xs text-gray-400 max-w-sm mx-auto mb-5">
                    Intenta buscando con otra palabra o revisa que esté escrita correctamente.
                  </p>
                  <button
                    type="button"
                    onClick={() => setSearchQuery("")}
                    className="px-6 py-2.5 rounded-xl bg-[#ffd025] text-black font-black uppercase text-xs tracking-wider hover:bg-[#ffe066] transition cursor-pointer shadow-lg shadow-[#ffd025]/20"
                  >
                    Ver catálogo completo
                  </button>
                </div>
              ) : (
                <div className="grid grid-cols-2 xs:grid-cols-2 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-7 2xl:grid-cols-8 gap-x-2 sm:gap-x-3.5 gap-y-4 sm:gap-y-6 items-stretch w-full">
                  {searchResults.map((product) => (
                    <div
                      key={product.id}
                      className="group flex flex-col justify-between h-full w-full"
                    >
                      <div>
                        <div className="relative w-full aspect-square overflow-hidden bg-black/40 mb-1.5 sm:mb-2">
                          {product.image ? (
                            <img
                              src={product.image}
                              alt={product.name}
                              loading="lazy"
                              decoding="async"
                              className="w-full h-full object-cover rounded-none group-hover:scale-105 transition-transform duration-300"
                              onError={(e) => {
                                (e.currentTarget as HTMLImageElement).src =
                                  "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                              }}
                            />
                          ) : (
                            <div className="w-full h-full flex flex-col items-center justify-center text-gray-500 bg-white/5 rounded-none">
                              <Package size={20} className="text-[#ffd025]/70" />
                              <span className="text-[8px] sm:text-[10px] mt-0.5 font-semibold uppercase">Fellas</span>
                            </div>
                          )}

                          {product.oferta && (
                            <div className="absolute top-0 left-0 z-10">
                              <span className="px-1.5 py-0.5 bg-red-600 text-white text-[8px] sm:text-[9px] font-black uppercase tracking-wider rounded-none shadow">
                                OFERTA
                              </span>
                            </div>
                          )}
                        </div>

                        <div className="h-3.5 sm:h-4 flex items-center mb-0.5 overflow-hidden">
                          {(product.subcategory || product.category || product.aisle) ? (
                            <span className="text-[7.5px] xs:text-[8px] sm:text-[9px] font-medium uppercase tracking-wider text-gray-400 truncate block w-full">
                              {product.subcategory || product.category || product.aisle}
                            </span>
                          ) : (
                            <span className="text-[7.5px] sm:text-[9px] font-medium uppercase tracking-wider text-transparent select-none">
                              -
                            </span>
                          )}
                        </div>

                        <h4
                          title={product.name}
                          className="text-[9.5px] xs:text-[10px] sm:text-[11.5px] md:text-[12px] font-semibold text-white leading-tight line-clamp-2 h-7 sm:h-8 md:h-8.5 block w-full group-hover:text-[#ffd025] transition-colors"
                        >
                          {product.name}
                        </h4>
                      </div>

                      <div className="mt-2 pt-1.5 border-t border-white/10 flex items-center justify-between gap-1">
                        <span className="text-[11px] sm:text-xs md:text-sm font-black text-[#ffd025] truncate">
                          ${Number(product.price || 0).toLocaleString("es-CL")}
                        </span>

                        <button
                          onClick={() => handleAddToCartClick(product)}
                          disabled={!isStoreOpen}
                          className="h-6 px-1.5 sm:px-2.5 bg-[#ffd025] hover:bg-[#e5b81a] text-black font-black text-[10px] sm:text-xs uppercase transition-all flex items-center justify-center gap-1 rounded-none hover:scale-105 active:scale-95 disabled:opacity-40 disabled:cursor-not-allowed shrink-0"
                          title="Añadir al carrito"
                          aria-label="Añadir al carrito"
                        >
                          <Plus size={12} strokeWidth={2.5} />
                          <span className="hidden sm:inline text-[10px]">Añadir</span>
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </section>
          ) : showDedicatedProductsPage && !settings.contingencyMode ? (
            <div className="animate-fade-in pb-4">
              <div className="w-full px-3 sm:px-6 md:px-8 mt-3 sm:mt-5">
                {/* Botón Volver a Inicio */}
                <div className="mb-2 flex items-center justify-between">
                  <button
                    onClick={() => {
                      setShowDedicatedProductsPage(false);
                      setDedicatedViewMode("catalog");
                      setActiveAisle("");
                      setActiveCategory("");
                      window.scrollTo({ top: 0, behavior: "smooth" });
                    }}
                    className="inline-flex items-center gap-1.5 text-xs font-bold text-gray-400 hover:text-white transition-colors py-1"
                  >
                    <ArrowLeft size={14} /> Volver a Portada
                  </button>
                  <span className="text-[10px] font-mono text-[#ffd025] uppercase tracking-wider">
                    {dedicatedViewMode === "oportunidades" ? "Pestaña Exclusiva Oportunidades" : dedicatedViewMode === "packs" ? "Pestaña Exclusiva Packs" : "Pestaña de Pasillos"}
                  </span>
                </div>

                {/* Banner exclusivo de sección (Solo imagen pura sin textos ni gradientes encima) */}
                <div className="relative w-full h-24 sm:h-40 md:h-60 lg:h-72 xl:h-80 overflow-hidden rounded-none border border-white/10 mb-4 bg-black select-none shadow-xl">
                  <img
                    src={
                      dedicatedViewMode === "oportunidades"
                        ? (settings.opportunitiesBannerImage || "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=1600&auto=format&fit=crop&q=80")
                        : dedicatedViewMode === "packs"
                        ? (settings.packsBannerImage || "https://images.unsplash.com/photo-1543007630-9710e4a00a20?w=1600&auto=format&fit=crop&q=80")
                        : getAisleBannerImage(activeAisle, activeAisle ? aislesData.find((a) => a.name.toLowerCase() === activeAisle.toLowerCase())?.bannerImage : settings.aislesBannerImage)
                    }
                    alt={
                      dedicatedViewMode === "oportunidades"
                        ? "Oportunidades"
                        : dedicatedViewMode === "packs"
                        ? "Packs"
                        : "Pasillos"
                    }
                    loading="lazy"
                    decoding="async"
                    className="w-full h-full object-cover rounded-none"
                  />
                </div>

                {/* Vista Dedicada de Productos segun Modo */}
                {(() => {
                  const isOpp = dedicatedViewMode === "oportunidades";
                  const isPacks = dedicatedViewMode === "packs";

                  let rawList: Product[] = [];
                  if (isOpp) {
                    const rawIds = settings.opportunitiesProductIds;
                    if (rawIds && Array.isArray(rawIds) && rawIds.length > 0) {
                      rawList = rawIds.map((id) => products.find((p) => p.id === id && !p.hidden)).filter(Boolean) as Product[];
                    }
                    if (rawList.length === 0) {
                      rawList = products.filter((p) => !p.hidden && (p.oferta || p.bestseller || (p.price && p.price < 12000)));
                    }
                  } else if (isPacks) {
                    const rawIds = settings.packsProductIds;
                    if (rawIds && Array.isArray(rawIds) && rawIds.length > 0) {
                      rawList = rawIds.map((id) => products.find((p) => p.id === id && !p.hidden)).filter(Boolean) as Product[];
                    }
                    if (rawList.length === 0) {
                      rawList = products.filter((p) => !p.hidden && (p.name.toLowerCase().includes("pack") || p.category?.toLowerCase().includes("pack") || p.aisle?.toLowerCase().includes("pack")));
                    }
                  } else {
                    rawList = products.filter((p) => !p.hidden);
                  }

                  // 1. Filtrar por pasillo/categoría activo del menú superior si existe
                  let processed = rawList.filter((p) => {
                    if (activeAisle && p.aisle !== activeAisle) return false;
                    if (activeCategory && p.category !== activeCategory) return false;
                    return true;
                  });

                  // 2. Extraer subcategorías únicas presentes
                  const availableSubcats = Array.from(
                    new Set(processed.map((p) => p.subcategory || p.category).filter(Boolean))
                  );

                  // 3. Filtrar por subcategoría elegida en el menú desplegable
                  if (dedicatedSubcatFilter !== "all") {
                    processed = processed.filter(
                      (p) => p.subcategory === dedicatedSubcatFilter || p.category === dedicatedSubcatFilter
                    );
                  }

                  // 4. Filtrar por búsqueda interna
                  if (dedicatedSearchQuery.trim().length > 0) {
                    const q = normalize(dedicatedSearchQuery.trim());
                    processed = processed.filter((p) =>
                      normalize(p.name).includes(q) ||
                      normalize(p.category || "").includes(q) ||
                      normalize(p.aisle || "").includes(q) ||
                      normalize(p.subcategory || "").includes(q)
                    );
                  }

                  // 5. Aplicar ordenamiento
                  if (dedicatedSortBy === "price_asc") {
                    processed.sort((a, b) => (a.price || 0) - (b.price || 0));
                  } else if (dedicatedSortBy === "price_desc") {
                    processed.sort((a, b) => (b.price || 0) - (a.price || 0));
                  } else if (dedicatedSortBy === "name_asc") {
                    processed.sort((a, b) => (a.name || "").localeCompare(b.name || "", "es"));
                  } else if (dedicatedSortBy === "name_desc") {
                    processed.sort((a, b) => (b.name || "").localeCompare(a.name || "", "es"));
                  }

                  return (
                    <div className="space-y-6 mb-12">
                      {/* BARRA DE BÚSQUEDA, SUBCATEGORÍAS Y ORDENAMIENTO (OPTIMIZADA PARA MÓVIL Y PC) */}
                      <div className="w-full border-y border-white/15 py-2.5 my-3 flex items-center justify-between gap-2 sm:gap-4 flex-nowrap">
                        {/* 1. Buscador: Siempre visible, toma todo el ancho disponible en celular */}
                        <div className="flex-1 min-w-0 flex items-center gap-1.5 sm:gap-2">
                          <Search size={16} className="text-[#ffd025] shrink-0 opacity-90" />
                          <input
                            type="text"
                            value={dedicatedSearchQuery}
                            onChange={(e) => setDedicatedSearchQuery(e.target.value)}
                            placeholder="BUSCAR..."
                            className="w-full bg-transparent text-white text-xs sm:text-sm font-bold uppercase placeholder-gray-500 focus:outline-none tracking-wider truncate"
                          />
                          {dedicatedSearchQuery && (
                            <button
                              type="button"
                              onClick={() => setDedicatedSearchQuery("")}
                              className="text-gray-400 hover:text-white p-0.5 shrink-0 transition-colors"
                              aria-label="Limpiar búsqueda"
                            >
                              <X size={14} />
                            </button>
                          )}
                        </div>

                        {/* Controles laterales: Iconos compactos en celular, completos en PC */}
                        <div className="flex items-center gap-1.5 sm:gap-5 shrink-0">
                          {/* 2. Menú / Filtro de Subcategorías */}
                          {availableSubcats.length > 0 && (
                            <div className="relative flex items-center shrink-0 border-l border-white/15 pl-1.5 sm:pl-4">
                              {/* Versión Celular: Solo icono compacto con select nativo */}
                              <div className="sm:hidden relative flex items-center justify-center p-1.5 text-[#ffd025] hover:text-[#ffe066] cursor-pointer" title="Filtrar por subcategoría">
                                <Layers size={17} />
                                {dedicatedSubcatFilter !== "all" && (
                                  <span className="absolute top-0 right-0 w-2 h-2 rounded-full bg-[#ffd025] ring-1 ring-black" />
                                )}
                                <select
                                  value={dedicatedSubcatFilter}
                                  onChange={(e) => setDedicatedSubcatFilter(e.target.value)}
                                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer text-base bg-black"
                                  aria-label="Filtrar por subcategoría"
                                >
                                  <option value="all">Todas las subcategorías</option>
                                  {availableSubcats.map((sub) => (
                                    <option key={sub} value={sub}>{sub}</option>
                                  ))}
                                </select>
                              </div>

                              {/* Versión PC: Selector con etiqueta de texto */}
                              <div className="hidden sm:flex items-center gap-1.5 text-xs">
                                <Layers size={14} className="text-[#ffd025] shrink-0" />
                                <span className="text-gray-400 uppercase text-[10px] font-bold tracking-wider">Subcategoría:</span>
                                <select
                                  value={dedicatedSubcatFilter}
                                  onChange={(e) => setDedicatedSubcatFilter(e.target.value)}
                                  className="bg-transparent text-white text-xs font-bold uppercase tracking-wider focus:outline-none cursor-pointer border-b border-white/20 hover:border-[#ffd025] pb-0.5 transition-colors"
                                  aria-label="Filtrar por subcategoría"
                                >
                                  <option value="all" className="bg-[#141422] text-white">Todas las subcategorías</option>
                                  {availableSubcats.map((sub) => (
                                    <option key={sub} value={sub} className="bg-[#141422] text-white">{sub}</option>
                                  ))}
                                </select>
                              </div>
                            </div>
                          )}

                          {/* 3. Botón / Menú de Ordenamiento */}
                          <div className="relative flex items-center shrink-0 border-l border-white/15 pl-1.5 sm:pl-4">
                            {/* Versión Celular: Solo icono compacto con select nativo */}
                            <div className="sm:hidden relative flex items-center justify-center p-1.5 text-[#ffd025] hover:text-[#ffe066] cursor-pointer" title="Ordenar productos">
                              <Filter size={17} />
                              {dedicatedSortBy !== "default" && (
                                <span className="absolute top-0 right-0 w-2 h-2 rounded-full bg-[#ffd025] ring-1 ring-black" />
                              )}
                              <select
                                value={dedicatedSortBy}
                                onChange={(e) => setDedicatedSortBy(e.target.value as any)}
                                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer text-base bg-black"
                                aria-label="Ordenar productos"
                              >
                                <option value="default">Destacados</option>
                                <option value="price_asc">Precio: Menor a Mayor ⬇️</option>
                                <option value="price_desc">Precio: Mayor a Menor ⬆️</option>
                                <option value="name_asc">Nombre: A ➔ Z</option>
                                <option value="name_desc">Nombre: Z ➔ A</option>
                              </select>
                            </div>

                            {/* Versión PC: Selector con etiqueta de texto */}
                            <div className="hidden sm:flex items-center gap-1.5 text-xs">
                              <Filter size={14} className="text-[#ffd025] shrink-0" />
                              <span className="text-gray-400 uppercase text-[10px] font-bold tracking-wider">Ordenar:</span>
                              <select
                                value={dedicatedSortBy}
                                onChange={(e) => setDedicatedSortBy(e.target.value as any)}
                                className="bg-transparent text-white text-xs font-bold uppercase tracking-wider focus:outline-none cursor-pointer border-b border-white/20 hover:border-[#ffd025] pb-0.5 transition-colors"
                                aria-label="Ordenar productos"
                              >
                                <option value="default" className="bg-[#141422] text-white">Destacados</option>
                                <option value="price_asc" className="bg-[#141422] text-white">Precio: Menor a Mayor ⬇️</option>
                                <option value="price_desc" className="bg-[#141422] text-white">Precio: Mayor a Menor ⬆️</option>
                                <option value="name_asc" className="bg-[#141422] text-white">Nombre: A ➔ Z</option>
                                <option value="name_desc" className="bg-[#141422] text-white">Nombre: Z ➔ A</option>
                              </select>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Lista horizontal rápida de subcategorías con líneas divisorias si existen */}
                      {availableSubcats.length > 1 && (
                        <div className="flex items-center gap-2 sm:gap-3 overflow-x-auto pb-1.5 -mt-2 border-b border-white/10 scrollbar-none">
                          <button
                            type="button"
                            onClick={() => setDedicatedSubcatFilter("all")}
                            className={`text-[10px] sm:text-xs uppercase tracking-wider whitespace-nowrap transition-colors py-0.5 ${
                              dedicatedSubcatFilter === "all"
                                ? "text-[#ffd025] font-black border-b-2 border-[#ffd025]"
                                : "text-gray-400 hover:text-white font-semibold"
                            }`}
                          >
                            Todas ({processed.length})
                          </button>
                          {availableSubcats.map((sub) => {
                            const subCount = processed.filter((p) => (p.subcategory || p.category) === sub).length;
                            return (
                              <Fragment key={sub}>
                                <span className="text-white/15 text-[10px] select-none">•</span>
                                <button
                                  type="button"
                                  onClick={() => setDedicatedSubcatFilter(sub)}
                                  className={`text-[10px] sm:text-xs uppercase tracking-wider whitespace-nowrap transition-colors py-0.5 ${
                                    dedicatedSubcatFilter === sub
                                      ? "text-[#ffd025] font-black border-b-2 border-[#ffd025]"
                                      : "text-gray-400 hover:text-white font-semibold"
                                  }`}
                                >
                                  {sub} ({subCount})
                                </button>
                              </Fragment>
                            );
                          })}
                        </div>
                      )}

                      {/* Header con conteo */}
                      <div className="flex items-center justify-between pb-2 border-b border-white/10">
                        <span className="text-xs sm:text-sm font-black uppercase text-[#ffd025] tracking-wider">
                          {isOpp
                            ? "Lista de Oportunidades"
                            : isPacks
                            ? "Lista de Packs & Combos"
                            : activeAisle
                            ? `Pasillo: ${activeAisle}`
                            : activeCategory
                            ? `Categoría: ${activeCategory}`
                            : "Todos los Productos"}
                        </span>
                        <span className="text-[10px] text-gray-400 font-bold uppercase">
                          {processed.length} {processed.length === 1 ? "producto" : "productos"}
                        </span>
                      </div>

                      {/* Grilla de Productos */}
                      {processed.length === 0 ? (
                        <div className="text-center py-12 border border-dashed border-white/10 rounded-2xl bg-white/[0.02]">
                          <p className="text-xs text-gray-400">No hay productos que coincidan con la búsqueda o filtros seleccionados.</p>
                        </div>
                      ) : (() => {
                        const len = processed.length;
                        const gridColsClass = len === 1
                          ? "md:grid-cols-1 w-full"
                          : len === 2
                          ? "md:grid-cols-2"
                          : len === 3
                          ? "md:grid-cols-3"
                          : len === 4
                          ? "md:grid-cols-4"
                          : len === 5
                          ? "md:grid-cols-5"
                          : len === 6
                          ? "md:grid-cols-6"
                          : len === 7
                          ? "md:grid-cols-7"
                          : "md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-7 2xl:grid-cols-8";
                        const mobileColsClass = len <= 4 ? "grid-cols-2" : "grid-cols-2 xs:grid-cols-3";
                        return (
                          <div className={`grid ${mobileColsClass} sm:grid-cols-4 gap-x-2 sm:gap-x-3.5 gap-y-4 sm:gap-y-6 items-stretch w-full ${gridColsClass}`}>
                            {processed.map((product) => (
                              <div
                                key={product.id}
                                className="group flex flex-col justify-between h-full w-full"
                              >
                                <div>
                                  <div className="relative w-full aspect-square overflow-hidden bg-black/40 mb-1.5 sm:mb-2">
                                    {product.image ? (
                                      <img
                                        src={product.image}
                                        alt={product.name}
                                        loading="lazy"
                                        decoding="async"
                                        className="w-full h-full object-cover rounded-none group-hover:scale-105 transition-transform duration-300"
                                        onError={(e) => {
                                          (e.currentTarget as HTMLImageElement).src =
                                            "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                        }}
                                      />
                                    ) : (
                                      <div className="w-full h-full flex flex-col items-center justify-center text-gray-500 bg-white/5 rounded-none">
                                        <Package size={20} className="text-[#ffd025]/70" />
                                        <span className="text-[8px] sm:text-[10px] mt-0.5 font-semibold uppercase">Fellas</span>
                                      </div>
                                    )}

                                    {(product.oferta || isOpp || isPacks) && (
                                      <div className="absolute top-0 left-0 z-10">
                                        <span className={`px-1.5 py-0.5 text-white text-[8px] sm:text-[9px] font-black uppercase tracking-wider rounded-none shadow ${isOpp ? "bg-amber-500" : isPacks ? "bg-purple-600" : "bg-red-600"}`}>
                                          {isOpp ? "OFERTA" : isPacks ? "PACK" : "PROMO"}
                                        </span>
                                      </div>
                                    )}
                                  </div>

                                  <div className="h-3.5 sm:h-4 flex items-center mb-0.5 overflow-hidden">
                                    {(product.subcategory || product.category || product.aisle) ? (
                                      <span className="text-[7.5px] xs:text-[8px] sm:text-[9px] font-medium uppercase tracking-wider text-gray-400 truncate block w-full">
                                        {product.subcategory || product.category || product.aisle}
                                      </span>
                                    ) : (
                                      <span className="text-[7.5px] sm:text-[9px] font-medium uppercase tracking-wider text-transparent select-none">
                                        -
                                      </span>
                                    )}
                                  </div>

                                  <h4
                                    title={product.name}
                                    className="text-[9.5px] xs:text-[10px] sm:text-[11.5px] md:text-[12px] font-semibold text-white leading-tight line-clamp-2 h-7 sm:h-8 md:h-8.5 block w-full group-hover:text-[#ffd025] transition-colors"
                                  >
                                    {product.name}
                                  </h4>
                                </div>

                                <div className="mt-2 pt-1.5 border-t border-white/10 flex items-center justify-between gap-1">
                                  <span className="text-[11px] sm:text-xs md:text-sm font-black text-[#ffd025] truncate">
                                    ${Number(product.price || 0).toLocaleString("es-CL")}
                                  </span>

                                  <button
                                    onClick={() => handleAddToCartClick(product)}
                                    disabled={!isStoreOpen}
                                    className="h-6 px-1.5 sm:px-2.5 bg-[#ffd025] hover:bg-[#e5b81a] text-black font-black text-[10px] sm:text-xs uppercase transition-all flex items-center justify-center gap-1 rounded-none hover:scale-105 active:scale-95 disabled:opacity-40 disabled:cursor-not-allowed shrink-0"
                                    title="Añadir al carrito"
                                    aria-label="Añadir al carrito"
                                  >
                                    <Plus size={12} strokeWidth={2.5} />
                                    <span className="hidden sm:inline text-[10px]">Añadir</span>
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        );
                      })()}
                    </div>
                  );
                })()}
              </div>
            </div>
          ) : (
            <>
              {/* BANNER PRINCIPAL O BANNER DE PASILLO SELECCIONADO */}
              {!searchQuery && (() => {
                const selectedAisleObj = activeAisle ? aislesData.find((a) => a.name === activeAisle) : null;
                const bannerToShow = selectedAisleObj?.bannerImage
                  ? [{ image: selectedAisleObj.bannerImage, title: selectedAisleObj.bannerTitle, description: selectedAisleObj.bannerSubtitle }]
                  : effectiveSlides;

                if (bannerToShow.length === 0) return null;

                return (
                  <div className="w-full px-3 sm:px-6 md:px-8 mt-3 sm:mt-5 mb-6 sm:mb-8">
                    <div
                      className="relative w-full h-36 sm:h-48 md:h-64 lg:h-[280px] xl:h-[320px] shadow-lg border border-white/10 bg-black select-none rounded-xl sm:rounded-2xl overflow-hidden"
                      style={{ transform: "translateZ(0)", WebkitMaskImage: "-webkit-radial-gradient(white, black)" }}
                    >
                      {bannerToShow.map((slide, i) => (
                        <div
                          key={i}
                          className="absolute inset-0 transition-opacity duration-700"
                          style={{ opacity: i === activeSlide ? 1 : 0, zIndex: i === activeSlide ? 2 : 1 }}
                        >
                          {slide.image && (
                            <img
                              loading={i === 0 ? "eager" : "lazy"}
                              decoding="async"
                              fetchPriority={i === 0 ? "high" : "auto"}
                              src={slide.image}
                              alt="Banner"
                              className="w-full h-full object-cover"
                            />
                          )}
                        </div>
                      ))}
                      {bannerToShow.length > 1 && !selectedAisleObj && (
                        <div className="absolute bottom-3 left-1/2 -translate-x-1/2 z-30 flex items-center gap-2">
                          {bannerToShow.map((_, i) => (
                            <button
                              key={i}
                              onClick={() => setActiveSlide(i)}
                              className={`rounded-full transition-all duration-300 ${
                                i === activeSlide
                                  ? "w-5 h-2 bg-[#ffd025]"
                                  : "w-2 h-2 bg-white/40 hover:bg-white/70"
                              }`}
                            />
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}

              {/* ESTADO DE FILTRO ACTIVO (CAMBIO DE PASILLO / CATEGORÍA DESDE EL ENCABEZADO) */}
              {(() => {
                const isFiltered = Boolean(activeAisle || activeCategory || navQuickFilter);

                const recommendedSection = promoProducts.length === 0 ? null : (
                  <section className="w-full mb-6 sm:mb-8">
                    <div className="w-full px-3 sm:px-6 md:px-8 mb-3 sm:mb-4 flex items-center gap-3 sm:gap-4">
                      <div className="flex-1 h-px bg-gradient-to-r from-transparent via-white/20 to-white/20" />
                      <span className="text-[10px] sm:text-xs md:text-sm font-black tracking-widest text-[#ffd025] uppercase shrink-0 px-1 select-none whitespace-nowrap drop-shadow">
                        #NUESTROSRECOMENDADOS
                      </span>
                      <div className="flex-1 h-px bg-gradient-to-l from-transparent via-white/20 to-white/20" />
                    </div>

                    <div className="w-full px-3 sm:px-6 md:px-8">
                      <div className="w-full h-16 sm:h-28 md:h-48 lg:h-60 xl:h-72 overflow-hidden rounded-none border border-white/10 mb-3 sm:mb-4 bg-black select-none">
                        <img
                          src={settings.promoBannerImage || "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=1600&auto=format&fit=crop&q=80"}
                          alt="Promociones Fellas"
                          loading="lazy"
                          decoding="async"
                          className="w-full h-full object-cover rounded-none"
                        />
                      </div>

                      {(() => {
                        const activePromoProducts = promoProducts.slice(0, 4);
                        const len = activePromoProducts.length;
                        const gridColsClass = len === 1
                          ? "md:grid-cols-1 w-full"
                          : len === 2
                          ? "md:grid-cols-2"
                          : len === 3
                          ? "md:grid-cols-3"
                          : "md:grid-cols-4 lg:grid-cols-4";
                        return (
                          <div>
                            <div className={`grid grid-cols-2 xs:grid-cols-2 sm:grid-cols-4 gap-x-2.5 sm:gap-x-3.5 gap-y-4 sm:gap-y-6 w-full items-stretch ${gridColsClass}`}>
                              {activePromoProducts.map((product) => (
                                <div
                                  key={product.id}
                                  className="group flex flex-col justify-between h-full w-full"
                                >
                                  <div>
                                    <div className="relative w-full aspect-square overflow-hidden bg-black/40 mb-1.5 sm:mb-2">
                                      {product.image ? (
                                        <img
                                          src={product.image}
                                          alt={product.name}
                                          loading="lazy"
                                          decoding="async"
                                          className="w-full h-full object-cover rounded-none group-hover:scale-105 transition-transform duration-300"
                                          onError={(e) => {
                                            (e.currentTarget as HTMLImageElement).src =
                                              "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                          }}
                                        />
                                      ) : (
                                        <div className="w-full h-full flex flex-col items-center justify-center text-gray-500 bg-white/5 rounded-none">
                                          <Package size={20} className="text-[#ffd025]/70" />
                                          <span className="text-[8px] sm:text-[10px] mt-0.5 font-semibold uppercase">Fellas</span>
                                        </div>
                                      )}

                                      <div className="absolute top-0 left-0 z-10">
                                        <span className="px-1.5 py-0.5 bg-red-600 text-white text-[8px] sm:text-[9px] font-black uppercase tracking-wider rounded-none shadow">
                                          PROMO
                                        </span>
                                      </div>
                                    </div>

                                    <div className="h-3.5 sm:h-4 flex items-center mb-0.5 overflow-hidden">
                                      {(product.aisle || product.category) ? (
                                        <span className="text-[7.5px] xs:text-[8px] sm:text-[9px] font-medium uppercase tracking-wider text-gray-400 truncate block w-full">
                                          {product.aisle || product.category}
                                        </span>
                                      ) : (
                                        <span className="text-[7.5px] sm:text-[9px] font-medium uppercase tracking-wider text-transparent select-none">
                                          -
                                        </span>
                                      )}
                                    </div>

                                    <h4
                                      title={product.name}
                                      className="text-[9.5px] xs:text-[10px] sm:text-[11.5px] md:text-[12px] font-semibold text-white leading-tight line-clamp-2 h-7 sm:h-8 md:h-8.5 block w-full group-hover:text-[#ffd025] transition-colors"
                                    >
                                      {product.name}
                                    </h4>
                                  </div>

                                  <div className="mt-2 pt-1.5 border-t border-white/10 flex items-center justify-between gap-1">
                                    <span className="text-[11px] sm:text-xs md:text-sm font-black text-[#ffd025] truncate">
                                      ${Number(product.price || 0).toLocaleString("es-CL")}
                                    </span>

                                    <button
                                      onClick={() => handleAddToCartClick(product)}
                                      disabled={!isStoreOpen}
                                      className="h-6 px-1.5 sm:px-2.5 bg-[#ffd025] hover:bg-[#e5b81a] text-black font-black text-[10px] sm:text-xs uppercase transition-all flex items-center justify-center gap-1 rounded-none hover:scale-105 active:scale-95 disabled:opacity-40 disabled:cursor-not-allowed shrink-0"
                                      title="Añadir al carrito"
                                      aria-label="Añadir al carrito"
                                    >
                                      <Plus size={12} strokeWidth={2.5} />
                                      <span className="hidden sm:inline text-[10px]">Añadir</span>
                                    </button>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  </section>
                );

                const showCollections = baseProducts.length > 0 || homeCollectionProducts.length > 0;

                const selectedAisleObj = activeAisle ? aislesData.find((a) => a.name.toLowerCase() === activeAisle.toLowerCase()) : null;
                const currentAisleBanner = activeAisle
                  ? getAisleBannerImage(activeAisle, selectedAisleObj?.bannerImage)
                  : (settings.aislesBannerImage || "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=1600&auto=format&fit=crop&q=80");

                const collectionsSection = !showCollections ? null : (
                  <section className="w-full mb-6 sm:mb-8">
                    <div className="w-full px-3 sm:px-6 md:px-8 mb-3 sm:mb-4 flex items-center gap-3 sm:gap-4">
                      <div className="flex-1 h-px bg-gradient-to-r from-transparent via-white/20 to-white/20" />
                      <span className="text-[10px] sm:text-xs md:text-sm font-black tracking-widest text-[#ffd025] uppercase shrink-0 px-1 select-none whitespace-nowrap drop-shadow">
                        #NUESTRASCOLECCIONES
                      </span>
                      <div className="flex-1 h-px bg-gradient-to-l from-transparent via-white/20 to-white/20" />
                      <button
                        onClick={() => {
                          setShowDedicatedProductsPage(true);
                          setDedicatedViewMode("catalog");
                          setActiveAisle("");
                          setActiveCategory("");
                          window.scrollTo({ top: 0, behavior: "smooth" });
                        }}
                        className="text-[10.5px] sm:text-xs font-semibold text-gray-400 hover:text-[#ffd025] uppercase tracking-wider transition-colors flex items-center gap-0.5 cursor-pointer shrink-0"
                      >
                        <span>Ver más</span>
                        <ChevronRight size={13} />
                      </button>
                    </div>

                    <div className="w-full px-3 sm:px-6 md:px-8">
                      <div className="w-full h-16 sm:h-28 md:h-48 lg:h-60 xl:h-72 overflow-hidden rounded-none border border-white/10 mb-3 sm:mb-4 bg-black select-none relative">
                        <AnimatePresence mode="wait">
                          <motion.img
                            key={currentAisleBanner}
                            src={currentAisleBanner}
                            alt={activeAisle ? `Pasillo ${activeAisle}` : "Pasillos"}
                            initial={{ opacity: 0.3 }}
                            animate={{ opacity: 1 }}
                            exit={{ opacity: 0.3 }}
                            transition={{ duration: 0.3 }}
                            loading="lazy"
                            decoding="async"
                            className="w-full h-full object-cover rounded-none"
                          />
                        </AnimatePresence>
                        {activeAisle && (
                          <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/20 to-transparent flex items-end p-3 sm:p-4">
                            <span className="text-[#ffd025] font-black text-xs sm:text-base uppercase tracking-wider drop-shadow">
                              Pasillo: {activeAisle}
                            </span>
                          </div>
                        )}
                      </div>

                      {/* DESGLOSE POR PASILLO */}
                      <div className="space-y-6 sm:space-y-8">
                        {activeAisles.map((aisleName) => {
                          const rawAisleProducts = groupedByAisle[aisleName] || [];
                          if (rawAisleProducts.length === 0) return null;
                          const aisleProducts = isFiltered ? rawAisleProducts : rawAisleProducts.slice(0, 4);

                          return (
                            <div key={aisleName} id={`aisle-${normalize(aisleName)}`} className="scroll-mt-24">
                              {/* GRILLA DE PRODUCTOS DEL PASILLO */}
                              <div className="grid grid-cols-2 xs:grid-cols-2 sm:grid-cols-4 gap-x-2.5 sm:gap-x-3.5 gap-y-4 sm:gap-y-6 items-stretch w-full">
                                {aisleProducts.map((product) => (
                                  <div
                                    key={product.id}
                                    className="group flex flex-col justify-between h-full w-full"
                                  >
                                    <div>
                                      <div className="relative w-full aspect-square overflow-hidden bg-black/40 mb-1.5 sm:mb-2">
                                        {product.image ? (
                                          <img
                                            src={product.image}
                                            alt={product.name}
                                            loading="lazy"
                                            decoding="async"
                                            className="w-full h-full object-cover rounded-none group-hover:scale-105 transition-transform duration-300"
                                            onError={(e) => {
                                              (e.currentTarget as HTMLImageElement).src =
                                                "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                            }}
                                          />
                                        ) : (
                                          <div className="w-full h-full flex flex-col items-center justify-center text-gray-500 bg-white/5 rounded-none">
                                            <Package size={20} className="text-[#ffd025]/70" />
                                            <span className="text-[8px] sm:text-[10px] mt-0.5 font-semibold uppercase">Fellas</span>
                                          </div>
                                        )}

                                        {product.oferta && (
                                          <div className="absolute top-0 left-0 z-10">
                                            <span className="px-1.5 py-0.5 bg-red-600 text-white text-[8px] sm:text-[9px] font-black uppercase tracking-wider rounded-none shadow">
                                              OFERTA
                                            </span>
                                          </div>
                                        )}
                                      </div>

                                      <div className="h-3.5 sm:h-4 flex items-center mb-0.5 overflow-hidden">
                                        {(product.subcategory || product.category || product.aisle) ? (
                                          <span className="text-[7.5px] xs:text-[8px] sm:text-[9px] font-medium uppercase tracking-wider text-gray-400 truncate block w-full">
                                            {product.subcategory || product.category || product.aisle}
                                          </span>
                                        ) : (
                                          <span className="text-[7.5px] sm:text-[9px] font-medium uppercase tracking-wider text-transparent select-none">
                                            -
                                          </span>
                                        )}
                                      </div>

                                      <h4
                                        title={product.name}
                                        className="text-[9.5px] xs:text-[10px] sm:text-[11.5px] md:text-[12px] font-semibold text-white leading-tight line-clamp-2 h-7 sm:h-8 md:h-8.5 block w-full group-hover:text-[#ffd025] transition-colors"
                                      >
                                        {product.name}
                                      </h4>
                                    </div>

                                    <div className="mt-2 pt-1.5 border-t border-white/10 flex items-center justify-between gap-1">
                                      <span className="text-[11px] sm:text-xs md:text-sm font-black text-[#ffd025] truncate">
                                        ${Number(product.price || 0).toLocaleString("es-CL")}
                                      </span>

                                      <button
                                        onClick={() => handleAddToCartClick(product)}
                                        disabled={!isStoreOpen}
                                        className="h-6 px-1.5 sm:px-2.5 bg-[#ffd025] hover:bg-[#e5b81a] text-black font-black text-[10px] sm:text-xs uppercase transition-all flex items-center justify-center gap-1 rounded-none hover:scale-105 active:scale-95 disabled:opacity-40 disabled:cursor-not-allowed shrink-0"
                                        title="Añadir al carrito"
                                        aria-label="Añadir al carrito"
                                      >
                                        <Plus size={12} strokeWidth={2.5} />
                                        <span className="hidden sm:inline text-[10px]">Añadir</span>
                                      </button>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {/* Botón Ver catálogo completo al final */}
                      {!isFiltered && (
                        <div className="w-full flex justify-center mt-8 sm:mt-10 pb-2">
                          <button
                            onClick={() => {
                              setShowDedicatedProductsPage(true);
                              setDedicatedViewMode("catalog");
                              setActiveAisle("");
                              setActiveCategory("");
                              window.scrollTo({ top: 0, behavior: "smooth" });
                            }}
                            className="px-8 py-3 bg-[#ffd025] hover:bg-[#e5b81a] text-black font-black text-xs sm:text-sm uppercase tracking-widest transition-all rounded-none hover:scale-105 active:scale-95 flex items-center justify-center gap-2 cursor-pointer shadow-lg shadow-[#ffd025]/10"
                          >
                            <span>Ver catálogo completo</span>
                          </button>
                        </div>
                      )}
                    </div>
                  </section>
                );

                // En estado inicial (Home, sin filtro de pasillo): #NUESTROSRECOMENDADOS siempre va PRIMERO, #NUESTRASCOLECCIONES SEGUNDO.
                // Al hacer clic en un pasillo específico: se intercambian de lugar animadamente de arriba a abajo.
                const sectionsOrder = isFiltered
                  ? ["collections", "recommended"]
                  : ["recommended", "collections"];

                return (
                  <div className="w-full space-y-6 sm:space-y-8">
                    <AnimatePresence mode="popLayout">
                      {sectionsOrder.map((sectionKey) => {
                        const content = sectionKey === "recommended" ? recommendedSection : collectionsSection;
                        if (!content) return null;
                        return (
                          <motion.div
                            key={sectionKey}
                            layout
                            initial={{ opacity: 0, y: 15 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -15 }}
                            transition={{
                              type: "spring",
                              stiffness: 280,
                              damping: 28,
                              mass: 0.8,
                            }}
                          >
                            {content}
                          </motion.div>
                        );
                      })}
                    </AnimatePresence>
                  </div>
                );
              })()}

              {/* SECCIÓN GALERÍA 4:5 SUELTA: 6 en PC, 2 en Celular, leve zoom, solo imagen sin textos ni botones */}
              <StoreGallery images={settings.galleryImages} />
            </>
          )}

          <footer className="mt-6 sm:mt-8 pt-6 pb-6 border-t border-white/10 bg-black/60 text-gray-400">
            <div className="w-full px-4 sm:px-8 md:px-10 lg:px-12">
              {/* Distribución exacta en 2 Columnas con línea divisoria central */}
              <div className="grid grid-cols-2 gap-3 sm:gap-8 items-start">
                {/* COLUMNA 1: Logo y Descripción de la Tienda */}
                <div className="flex flex-col items-start gap-2.5 pr-3 sm:pr-8 border-r border-white/15 h-full">
                  <div className="flex items-center justify-start min-h-[32px] sm:min-h-[40px]">
                    {settings.footerLogo || settings.logo ? (
                      <img
                        src={settings.footerLogo || settings.logo}
                        alt={settings.pageTitle || "Tienda"}
                        className="h-8 sm:h-11 w-auto max-w-[120px] sm:max-w-[160px] object-contain"
                      />
                    ) : (
                      <h3 className="text-xs sm:text-base font-black tracking-wider text-[#ffd025] uppercase">
                        {settings.pageTitle || "FELLA'S MARKET"}
                      </h3>
                    )}
                  </div>
                  <p className="text-[10px] sm:text-xs text-gray-300 leading-snug line-clamp-2 max-w-sm m-0">
                    {settings.footerDescription || settings.bannerDescription || "Tu botillería y minimarket de confianza. Cervezas heladas, destilados, snacks y delivery rápido."}
                  </p>
                </div>

                {/* COLUMNA 2: Redes Sociales y justo debajo los Pasillos de la página */}
                <div className="flex flex-col items-end gap-3 text-right pl-2 sm:pl-4">
                  {/* Redes Sociales */}
                  <div className="flex items-center justify-end gap-2">
                    <a
                      href={settings.socialInstagram || "https://instagram.com"}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="w-7 h-7 sm:w-8 sm:h-8 rounded-full bg-white/10 hover:bg-[#ffd025] hover:text-black border border-white/15 flex items-center justify-center text-gray-300 transition-all shadow shrink-0"
                      title="Instagram"
                      aria-label="Instagram"
                    >
                      <Instagram size={13} className="sm:w-4 sm:h-4" />
                    </a>
                    <a
                      href={settings.socialFacebook || "https://facebook.com"}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="w-7 h-7 sm:w-8 sm:h-8 rounded-full bg-white/10 hover:bg-[#ffd025] hover:text-black border border-white/15 flex items-center justify-center text-gray-300 transition-all shadow shrink-0"
                      title="Facebook"
                      aria-label="Facebook"
                    >
                      <Facebook size={13} className="sm:w-4 sm:h-4" />
                    </a>
                    <a
                      href={
                        settings.socialWhatsapp || settings.whatsapp || settings.contactPhone
                          ? `https://api.whatsapp.com/send?phone=${(settings.socialWhatsapp || settings.whatsapp || settings.contactPhone || "").replace(/\D/g, "")}`
                          : "https://whatsapp.com"
                      }
                      target="_blank"
                      rel="noopener noreferrer"
                      className="w-7 h-7 sm:w-8 sm:h-8 rounded-full bg-white/10 hover:bg-[#25D366] hover:text-white border border-white/15 flex items-center justify-center text-gray-300 transition-all shadow shrink-0"
                      title="WhatsApp"
                      aria-label="WhatsApp"
                    >
                      <MessageCircle size={13} className="sm:w-4 sm:h-4" />
                    </a>
                  </div>

                  {/* Justo debajo los Pasillos de la página */}
                  <div className="w-full flex flex-wrap justify-end gap-x-2.5 gap-y-1">
                    <button
                      onClick={() => {
                        setActiveCategory("");
                        setActiveAisle("");
                        setNavQuickFilter("");
                        window.scrollTo({ top: 0, behavior: "smooth" });
                      }}
                      className={`text-[9.5px] sm:text-[11px] uppercase tracking-wider transition-colors ${
                        !activeCategory && !activeAisle && !navQuickFilter
                          ? "text-[#ffd025] font-black"
                          : "text-gray-400 hover:text-white font-medium"
                      }`}
                    >
                      Todos
                    </button>
                    {(aisles.length > 0 ? aisles : categories)
                      .filter((item) => {
                        const isAisle = aisles.includes(item);
                        const count = baseProducts.filter((p) =>
                          isAisle ? p.aisle === item : p.category === item
                        ).length;
                        return count > 0;
                      })
                      .slice(0, 6)
                      .map((item) => {
                      const isAisleSelected = activeAisle === item;
                      const isCategorySelected = activeCategory === item;
                      const isSelected = isAisleSelected || isCategorySelected;
                      return (
                        <button
                          key={item}
                          onClick={() => {
                            setNavQuickFilter("");
                            if (aisles.includes(item)) {
                              setActiveCategory("");
                              setActiveAisle(item);
                            } else {
                              setActiveAisle("");
                              setActiveCategory(item);
                            }
                            window.scrollTo({ top: 0, behavior: "smooth" });
                          }}
                          className={`text-[9.5px] sm:text-[11px] uppercase tracking-wider transition-colors ${
                            isSelected
                              ? "text-[#ffd025] font-black"
                              : "text-gray-400 hover:text-white font-medium"
                          }`}
                        >
                          {item}
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Derechos de Autor y Créditos de Agencia / Desarrollador (Una sola línea breve) */}
              <div className="mt-5 pt-3 border-t border-white/10 flex items-center justify-between text-[9.5px] sm:text-[11px] text-gray-400 uppercase tracking-wider">
                <span className="font-semibold truncate">
                  © {settings.pageTitle || "FELLA'S MARKET"}
                </span>
                <div className="flex items-center gap-1.5 shrink-0 ml-3">
                  <span className="text-gray-500 font-medium">Por</span>
                  {settings.agencyLogo ? (
                    <img
                      src={settings.agencyLogo}
                      alt={settings.agencyName || "Logo Agencia"}
                      className="h-3.5 sm:h-4 object-contain opacity-85 hover:opacity-100 transition-opacity"
                    />
                  ) : (
                    <span className="text-[#ffd025] font-black tracking-wider">
                      {settings.agencyName || "Fella's Market"}
                    </span>
                  )}
                </div>
              </div>
            </div>
          </footer>
        </main>
        </div>
      )}

      {checkoutOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4">
          <div className="absolute inset-0 bg-black/80 backdrop-blur-sm" onClick={() => setCheckoutOpen(false)} />
          <div className="relative w-full max-w-sm bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl shadow-2xl animate-fade-in overflow-hidden">
            {/* Header */}
            <div className="flex items-center justify-between px-5 pt-5 pb-3">
              <div>
                <p className="text-[10px] font-bold text-gray-500 uppercase tracking-widest mb-0.5">
                  Paso {checkoutStep} de 3
                </p>
                <h3 className="text-base font-black uppercase tracking-wide">
                  {checkoutStep === 1 && "¿Cómo lo recibís?"}
                  {checkoutStep === 2 && "Tus datos"}
                  {checkoutStep === 3 && "Resumen y envío"}
                </h3>
              </div>
              <button type="button" onClick={() => setCheckoutOpen(false)} className="text-gray-500 hover:text-white transition-colors ml-4 shrink-0">
                <X size={20} />
              </button>
            </div>
            {/* Progress bar */}
            <div className="px-5 pb-4">
              <div className="flex gap-1.5">
                {[1, 2, 3].map(s => (
                  <div key={s} className={`h-1 flex-1 rounded-full transition-all duration-300 ${s <= checkoutStep ? "bg-[#ffd025]" : "bg-gray-800"}`} />
                ))}
              </div>
            </div>

            <form onSubmit={submitCheckout}>
              <div className="px-5 pb-5">

                {/* ── PASO 1: Tipo de entrega ── */}
                {checkoutStep === 1 && (
                  <div className="space-y-4 animate-fade-in">
                    <div className="grid grid-cols-2 gap-3">
                      {(["retiro", "delivery"] as const).map(type => (
                        <button
                          key={type}
                          type="button"
                          onClick={() => setCheckoutForm({ ...checkoutForm, deliveryType: type })}
                          className={`flex flex-col items-center justify-center gap-2 py-5 px-3 rounded-2xl border-2 font-bold text-sm transition-all ${
                            checkoutForm.deliveryType === type
                              ? "bg-[#ffd025]/10 border-[#ffd025] text-[#ffd025]"
                              : "bg-[#141414] border-gray-700 text-gray-400 hover:border-[#ffd025]/40"
                          }`}
                        >
                          <span className="text-3xl">{type === "retiro" ? "🏪" : "🛵"}</span>
                          <span>{type === "retiro" ? "Retiro en local" : "Delivery"}</span>
                        </button>
                      ))}
                    </div>
                    {checkoutForm.deliveryType === "delivery" && !deliveryMinimumMet && (
                      <p className="text-[11px] text-amber-400 bg-amber-950/40 border border-amber-600/30 rounded-xl px-3 py-2.5 leading-snug">
                        ⚠️ Mínimo para delivery: <span className="font-black">${deliveryMinimum.toLocaleString("es-CL")}</span>. Agrega más productos.
                      </p>
                    )}
                    {checkoutForm.deliveryType === "delivery" && hasLocations && (
                      <div className="space-y-2">
                        <p className="text-[11px] font-bold text-gray-400 uppercase">Zona de envío</p>
                        <div className="space-y-2">
                          {deliveryLocations.filter(l => l.active).map(loc => (
                            <button
                              key={loc.id}
                              type="button"
                              onClick={() => setSelectedLocation(loc)}
                              className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border-2 font-semibold text-sm transition-all ${
                                selectedLocation?.id === loc.id
                                  ? "bg-[#ffd025]/10 border-[#ffd025] text-[#ffd025]"
                                  : "bg-[#141414] border-gray-700 text-gray-300 hover:border-[#ffd025]/40"
                              }`}
                            >
                              <span>📍 {loc.name}</span>
                              <span className="font-black">${loc.price.toLocaleString("es-CL")}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                    <div className="flex items-center gap-2 rounded-xl px-3 py-2.5 border border-[#ffd025]/20" style={{ background: "linear-gradient(135deg,rgba(255,208,37,0.08),rgba(255,138,0,0.06))" }}>
                      <span className="text-base">💳</span>
                      <p className="text-[11px] text-[#ffd025]/80 font-semibold leading-snug">
                        Pago por <span className="font-black text-[#ffd025]">transferencia</span>. Los datos se coordinan por WhatsApp.
                      </p>
                    </div>
                    <button
                      type="button"
                      disabled={
                        (checkoutForm.deliveryType === "delivery" && !deliveryMinimumMet) ||
                        (checkoutForm.deliveryType === "delivery" && hasLocations && !selectedLocation)
                      }
                      onClick={() => setCheckoutStep(customer ? 3 : 2)}
                      className="w-full py-3.5 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      Siguiente
                    </button>
                  </div>
                )}

                {/* ── PASO 2: Nombre y teléfono ── */}
                {checkoutStep === 2 && (
                  <div className="space-y-4 animate-fade-in">
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Nombre *</label>
                      <input
                        type="text"
                        required
                        autoFocus
                        value={checkoutForm.customerName}
                        onChange={e => setCheckoutForm({ ...checkoutForm, customerName: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                        placeholder="Tu nombre"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Teléfono *</label>
                      <input
                        type="tel"
                        required
                        value={checkoutForm.phone}
                        onChange={e => setCheckoutForm({ ...checkoutForm, phone: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                        placeholder="+56 9..."
                      />
                    </div>
                    <div className="flex gap-2 pt-1">
                      <button type="button" onClick={() => setCheckoutStep(1)} className="px-5 py-3 bg-[#141414] border border-gray-700 text-gray-400 rounded-2xl font-bold text-sm hover:border-gray-500 transition-all">
                        Atrás
                      </button>
                      <button
                        type="button"
                        disabled={!checkoutForm.customerName.trim() || !checkoutForm.phone.trim()}
                        onClick={() => setCheckoutStep(3)}
                        className="flex-1 py-3 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        Siguiente
                      </button>
                    </div>
                  </div>
                )}

                {/* ── PASO 3: Dirección, comentario, descuento y total ── */}
                {checkoutStep === 3 && (
                  <div className="space-y-3 animate-fade-in">
                    {/* Logged-in customer info chip */}
                    {customer && (
                      <div className="flex items-center gap-2.5 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-3 py-2.5">
                        <div className="w-7 h-7 rounded-full bg-[#ffd025]/15 flex items-center justify-center shrink-0">
                          <User size={14} className="text-[#ffd025]" />
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-bold text-white leading-tight truncate">{customer.name}</p>
                          <p className="text-xs text-gray-500 truncate">{customer.phone || customer.email}</p>
                        </div>
                      </div>
                    )}
                    {/* Phone field for logged-in customers without a saved phone */}
                    {customer && !checkoutForm.phone.trim() && (
                      <div>
                        <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Teléfono *</label>
                        <input
                          type="tel"
                          required
                          value={checkoutForm.phone}
                          onChange={e => setCheckoutForm({ ...checkoutForm, phone: e.target.value })}
                          className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                          placeholder="+56 9..."
                        />
                      </div>
                    )}
                    {isDelivery && (
                      <div>
                        <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Dirección de entrega *</label>
                        <input
                          type="text"
                          required={isDelivery}
                          value={checkoutForm.address}
                          onChange={e => setCheckoutForm({ ...checkoutForm, address: e.target.value })}
                          className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                          placeholder="Calle, número, comuna"
                        />
                      </div>
                    )}
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Comentarios (opcional)</label>
                      <textarea
                        rows={2}
                        value={checkoutForm.notes}
                        onChange={e => setCheckoutForm({ ...checkoutForm, notes: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025] resize-none"
                        placeholder="Instrucciones especiales..."
                      />
                    </div>
                    {/* Discount code */}
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5 flex items-center gap-1.5">
                        <Tag size={11} /> Código de descuento
                      </label>
                      {!checkoutForm.phone.trim() ? (
                        <p className="text-[11px] text-gray-500 bg-[#141414] border border-gray-800 rounded-xl px-3 py-2.5">
                          Ingresa tu teléfono para aplicar descuentos
                        </p>
                      ) : appliedDiscount ? (
                        <div className="flex items-center justify-between bg-green-900/30 border border-green-500/30 rounded-xl px-3 py-2">
                          <div className="flex items-center gap-2">
                            <CheckCircle size={13} className="text-green-400 shrink-0" />
                            <span className="text-green-400 text-sm font-bold">{appliedDiscount.label}</span>
                          </div>
                          <button type="button" onClick={() => { setAppliedDiscount(null); setDiscountCode(""); }} className="text-gray-500 hover:text-gray-300 text-xs ml-2 shrink-0">Quitar</button>
                        </div>
                      ) : (
                        <div className="flex gap-2 items-center">
                          <input
                            type="text"
                            value={discountCode}
                            onChange={e => { setDiscountCode(e.target.value.toUpperCase()); setDiscountError(""); }}
                            className="min-w-0 flex-1 bg-[#141414] border border-[#ffd025]/20 rounded-xl py-2 px-3 text-white text-sm uppercase tracking-widest focus:outline-none focus:border-[#ffd025] placeholder-gray-600"
                            placeholder="CÓDIGO"
                          />
                          <button type="button" onClick={applyDiscount} disabled={applyingDiscount || !discountCode.trim()}
                            className="px-3 py-2 bg-[#ffd025]/10 border border-[#ffd025]/20 text-[#ffd025] rounded-xl text-xs font-bold hover:bg-[#ffd025]/20 transition-colors disabled:opacity-40 shrink-0 whitespace-nowrap">
                            {applyingDiscount ? "…" : "Aplicar"}
                          </button>
                        </div>
                      )}
                      {discountError && <p className="text-red-400 text-xs mt-1">{discountError}</p>}
                    </div>
                    {/* Total */}
                    <div className="rounded-2xl bg-[#141414] border border-[#ffd025]/15 px-4 py-3 space-y-1.5">
                      <div className="flex justify-between text-sm text-gray-400">
                        <span>Subtotal</span>
                        <span>${cartTotal.toLocaleString("es-CL")}</span>
                      </div>
                      {isDelivery && (
                        <>
                          <div className="flex justify-between text-sm text-gray-400">
                            <span>Delivery</span>
                            <span>${DELIVERY_COST.toLocaleString("es-CL")}</span>
                          </div>
                          <p className="text-[10px] text-amber-400/70 leading-snug">🏘️ Tarifa fija dentro de Alerce.</p>
                        </>
                      )}
                      {discountSavings > 0 && (
                        <div className="flex justify-between text-sm text-green-400">
                          <span>Descuento ({appliedDiscount?.code})</span>
                          <span>−${discountSavings.toLocaleString("es-CL")}</span>
                        </div>
                      )}
                      <div className="flex justify-between font-black text-base text-[#ffd025] pt-1.5 border-t border-[#ffd025]/15">
                        <span>Total</span>
                        <span>${finalTotal.toLocaleString("es-CL")}</span>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <button type="button" onClick={() => setCheckoutStep(customer ? 1 : 2)} className="px-5 py-3 bg-[#141414] border border-gray-700 text-gray-400 rounded-2xl font-bold text-sm hover:border-gray-500 transition-all shrink-0">
                        Atrás
                      </button>
                      <button
                        type="submit"
                        disabled={createOrderMut.isPending || !isStoreOpen || !deliveryMinimumMet || (isDelivery && !checkoutForm.address.trim())}
                        className="flex-1 py-3 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all flex items-center justify-center gap-1.5 shadow-[0_0_20px_rgba(255,208,37,0.25)] disabled:opacity-40 disabled:cursor-not-allowed disabled:shadow-none"
                      >
                        <Phone size={14} /> Confirmar pedido
                      </button>
                    </div>
                  </div>
                )}

              </div>
            </form>
          </div>
        </div>
      )}

      {/* Customer auth modal */}
      {showAuthModal && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4">
          <div className="absolute inset-0 bg-black/80 backdrop-blur-sm" onClick={() => { setShowAuthModal(false); setAuthError(""); }} />
          <div className="relative w-full max-w-sm bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl p-6 shadow-2xl animate-fade-in">
            <div className="flex justify-between items-center mb-5">
              <h2 className="text-xl font-black uppercase">
                {authMode === "login" ? "Iniciar sesión" : "Crear cuenta"}
              </h2>
              <button onClick={() => { setShowAuthModal(false); setAuthError(""); }} className="text-gray-500 hover:text-white transition-colors">
                <X size={20} />
              </button>
            </div>
            <form onSubmit={handleCustomerAuth} className="space-y-3">
              {authMode === "register" && (
                <div>
                  <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Nombre</label>
                  <input type="text" required value={authFields.name} onChange={e => setAuthFields(f => ({ ...f, name: e.target.value }))}
                    className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="Tu nombre" />
                </div>
              )}
              <div>
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                  {authMode === "login" ? "Correo o usuario" : "Correo electrónico"}
                </label>
                <input
                  type="text"
                  required
                  value={authFields.email}
                  onChange={e => setAuthFields(f => ({ ...f, email: e.target.value }))}
                  className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                  placeholder={authMode === "login" ? "Correo o usuario (ej: admin, delivery)" : "correo@ejemplo.com"}
                />
              </div>
              {authMode === "register" && (
                <div>
                  <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Teléfono (opcional)</label>
                  <input type="tel" value={authFields.phone} onChange={e => setAuthFields(f => ({ ...f, phone: e.target.value }))}
                    className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="+56 9..." />
                </div>
              )}
              <div>
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Contraseña</label>
                <input type="password" required minLength={authMode === "register" ? 6 : 1} value={authFields.password} onChange={e => setAuthFields(f => ({ ...f, password: e.target.value }))}
                  className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder={authMode === "register" ? "Mínimo 6 caracteres" : "••••••"} />
              </div>
              {authError && <p className="text-red-400 text-sm text-center py-1">{authError}</p>}
              <button type="submit" disabled={authLoading}
                className="w-full py-3.5 bg-[#ffd025] text-[#141414] rounded-2xl font-black uppercase tracking-wider hover:bg-[#e5b81a] transition-all disabled:opacity-50 flex items-center justify-center gap-2 mt-1">
                {authLoading ? <Loader2 size={20} className="animate-spin" /> : (authMode === "login" ? "Ingresar" : "Crear cuenta")}
              </button>
            </form>
            <p className="text-center text-sm text-gray-500 mt-4">
              {authMode === "login" ? "¿No tienes cuenta?" : "¿Ya tienes cuenta?"}
              <button onClick={() => { setAuthMode(authMode === "login" ? "register" : "login"); setAuthError(""); }} className="text-[#ffd025] font-bold ml-1 hover:underline">
                {authMode === "login" ? "Crear cuenta" : "Iniciar sesión"}
              </button>
            </p>
          </div>
        </div>
      )}

      {/* Customer account panel */}
      {showAccountPanel && customer && (
        <div className="fixed inset-0 z-[65]" onClick={() => setShowAccountPanel(false)}>
          <div className="absolute top-[72px] right-4 w-80 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-2xl shadow-2xl overflow-hidden animate-fade-in" onClick={e => e.stopPropagation()}>
            <div className="p-4 border-b border-gray-800">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-[#ffd025]/20 rounded-full flex items-center justify-center shrink-0">
                  <User size={20} className="text-[#ffd025]" />
                </div>
                <div className="min-w-0">
                  <p className="font-bold text-sm truncate">{customer.name}</p>
                  <p className="text-gray-400 text-xs truncate">{customer.email}</p>
                </div>
              </div>
            </div>
            <div className="p-4">
              <div className="flex items-center justify-between mb-2.5">
                <h3 className="text-xs font-bold text-gray-400 uppercase tracking-wider">Carritos guardados</h3>
                {cart.length > 0 && (
                  <button onClick={saveCurrentCart} className="text-xs text-[#ffd025]/80 hover:text-[#ffd025] font-bold transition-colors flex items-center gap-1">
                    <Bookmark size={11} /> Guardar actual
                  </button>
                )}
              </div>
              {savedCarts.length === 0 ? (
                <p className="text-gray-600 text-xs text-center py-4">No tienes carritos guardados.</p>
              ) : (
                <div className="space-y-2 max-h-52 overflow-y-auto">
                  {savedCarts.map(c => (
                    <div key={c.id} className="flex items-center justify-between bg-[#141414] rounded-xl px-3 py-2.5">
                      <div className="min-w-0 mr-2">
                        <p className="text-sm font-bold truncate">{c.name}</p>
                        <p className="text-xs text-gray-500">{c.items.length} producto{c.items.length !== 1 ? "s" : ""}</p>
                      </div>
                      <div className="flex gap-1 shrink-0">
                        <button onClick={() => loadSavedCart(c)} className="text-xs px-2.5 py-1.5 bg-[#ffd025]/10 text-[#ffd025] rounded-lg hover:bg-[#ffd025]/20 transition-colors font-bold">
                          Cargar
                        </button>
                        <button onClick={() => deleteSavedCart(c.id)} className="p-1.5 text-red-400/50 hover:text-red-400 rounded-lg hover:bg-red-900/20 transition-colors">
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div className="p-3 border-t border-gray-800">
              <button onClick={handleCustomerLogout} className="w-full flex items-center justify-center gap-2 py-2.5 text-sm text-red-400 hover:text-red-300 hover:bg-red-900/20 rounded-xl transition-colors font-bold">
                <LogOut size={15} /> Cerrar sesión
              </button>
            </div>
          </div>
        </div>
      )}

      {view === "admin-login" && (
        <div className="min-h-screen flex items-center justify-center p-4 bg-[#141414]">
          <div className="bg-[#1a1a1a] p-8 rounded-3xl w-full max-w-sm border border-[#ffd025]/20 shadow-2xl text-center">
            <div className="w-20 h-20 bg-[#ffd025]/10 rounded-full flex items-center justify-center mx-auto mb-6">
              <SettingsIcon className="w-10 h-10 text-[#ffd025]" />
            </div>
            <h2 className="text-2xl font-black uppercase mb-2">Administración</h2>
            <p className="text-gray-400 text-sm mb-6">Ingresa tu contraseña para acceder.</p>
            <div className="mb-6">
              <input
                type="password"
                placeholder="Contraseña"
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
                className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-4 text-center text-white text-lg tracking-widest focus:outline-none focus:border-[#ffd025] transition-colors"
              />
            </div>
            <button
              onClick={handleLogin}
              className="w-full py-4 bg-[#ffd025] text-[#141414] rounded-xl font-bold uppercase hover:bg-[#e5b81a] transition-colors mb-4 shadow-lg shadow-[#ffd025]/20"
            >
              Ingresar al Panel
            </button>
            <button
              onClick={() => {
                setView("client");
                setPasswordInput("");
              }}
              className="w-full py-3 bg-transparent text-gray-500 font-bold uppercase hover:text-white transition-colors text-sm"
            >
              Volver a la tienda
            </button>
          </div>
        </div>
      )}

      {view === "admin" && (
        <div className="min-h-screen bg-[#111113] text-white flex flex-col md:flex-row font-sans">
          {/* Desktop Left Sidebar (Fixed / Sticky) */}
          <aside className="hidden md:flex md:w-64 lg:w-72 bg-[#17171c] border-r border-white/10 flex-col shrink-0 min-h-screen sticky top-0 h-screen overflow-y-auto custom-admin-scrollbar">
            {/* Sidebar Header / Brand */}
            <div className="p-5 border-b border-white/10 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-[#ffd025] to-[#d4a810] text-[#141414] font-black flex items-center justify-center shadow-lg shadow-[#ffd025]/20 shrink-0">
                  F
                </div>
                <div className="min-w-0">
                  <h2 className="font-black text-white text-base uppercase tracking-wider truncate">FELLAS</h2>
                  <p className="text-[11px] text-gray-400 truncate">Panel Autoadministrable</p>
                </div>
              </div>
            </div>

            {/* Role Badge */}
            <div className="px-5 py-3 bg-[#111113]/60 border-b border-white/5 flex items-center justify-between text-xs">
              <span className="text-gray-400">Rol activo:</span>
              <span className="font-bold text-[#ffd025] px-2.5 py-0.5 rounded-full bg-[#ffd025]/10 border border-[#ffd025]/20 uppercase text-[10px]">
                {adminRole === "delivery" ? "🛵 Delivery" : "👑 Admin Total"}
              </span>
            </div>

            {/* Navigation Links */}
            <nav className="flex-1 p-3 space-y-1 overflow-y-auto custom-admin-scrollbar">
              <div className="px-3 py-2 text-[10px] font-black text-gray-500 uppercase tracking-widest">
                Menú Principal
              </div>
              {[
                { id: "products", label: "Productos", icon: Package, desc: "Catálogo, fotos y stock" },
                { id: "media", label: "Archivo de Imágenes", icon: FolderOpen, desc: "Subida en grupo y galería" },
                { id: "classifications", label: "Clasificaciones", icon: Tag, desc: "Categorías y pasillos" },
                { id: "contingency", label: "Tienda Contingencia", icon: AlertTriangle, desc: "Pasillos y catálogo reducido" },
                { id: "orders", label: "Pedidos", icon: ClipboardList, desc: "Comandas en vivo" },
                { id: "stats", label: "Estadísticas", icon: TrendingUp, desc: "Reportes de ventas" },
                { id: "customers", label: "Clientes", icon: Users, desc: "Base de datos" },
                { id: "social", label: "Community", icon: Share2, desc: "Redes y avisos" },
                { id: "settings", label: "Ajustes Generales", icon: SettingsIcon, desc: "Configuración tienda" },
              ]
                .filter((item) => adminRole === "delivery" ? item.id === "orders" : true)
                .map((item) => {
                  const Icon = item.icon;
                  const isActive = adminTab === item.id;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => setAdminTab(item.id as any)}
                      className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl transition-all group text-left ${
                        isActive
                          ? "bg-[#ffd025] text-[#141414] font-bold shadow-lg shadow-[#ffd025]/20"
                          : "text-gray-400 hover:text-white hover:bg-white/5 font-medium"
                      }`}
                    >
                      <Icon size={18} className={`shrink-0 ${isActive ? "text-[#141414]" : "text-gray-400 group-hover:text-[#ffd025]"}`} />
                      <div className="min-w-0 flex-1">
                        <div className="text-xs uppercase tracking-wide truncate font-bold">{item.label}</div>
                        <div className={`text-[10px] truncate ${isActive ? "text-[#141414]/70" : "text-gray-500"}`}>{item.desc}</div>
                      </div>
                      {isActive && <ChevronRight size={14} className="text-[#141414] shrink-0" />}
                    </button>
                  );
                })}
            </nav>

            {/* Sidebar Footer Actions */}
            <div className="p-4 border-t border-white/10 space-y-2 bg-[#141418]">
              <button
                type="button"
                onClick={() => setView("client")}
                className="w-full flex items-center justify-center gap-2 px-4 py-2.5 bg-white/5 hover:bg-[#ffd025] hover:text-[#141414] text-gray-300 rounded-xl font-bold transition-all text-xs uppercase border border-white/10"
              >
                <Eye size={15} /> Ver Tienda Pública
              </button>
              <button
                type="button"
                onClick={() => { setView("client"); setAdminRole(null); }}
                className="w-full flex items-center justify-center gap-2 px-4 py-2 text-red-400 hover:text-red-300 hover:bg-red-500/10 rounded-xl font-bold transition-colors text-xs uppercase"
              >
                <LogOut size={14} /> Cerrar Sesión
              </button>
            </div>
          </aside>

          {/* Mobile Top Navigation Bar */}
          <div className="md:hidden bg-[#17171c] border-b border-white/10 p-3 sticky top-0 z-40 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={() => setAdminMobileMenuOpen(!adminMobileMenuOpen)}
                className="p-2 text-white bg-white/5 rounded-xl border border-white/10 hover:bg-white/10 transition-colors"
              >
                <Menu size={20} />
              </button>
              <div className="flex items-center gap-2">
                <span className="font-black text-[#ffd025] uppercase tracking-wider text-sm">FELLAS ADMIN</span>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setView("client")}
              className="px-3 py-1.5 bg-[#ffd025] text-[#141414] rounded-lg font-bold text-xs uppercase flex items-center gap-1.5 shadow"
            >
              <Eye size={14} /> Tienda
            </button>
          </div>

          {/* Mobile Slide-over Drawer Menu */}
          {adminMobileMenuOpen && (
            <div className="fixed inset-0 z-50 md:hidden flex">
              <div className="fixed inset-0 bg-black/80 backdrop-blur-sm" onClick={() => setAdminMobileMenuOpen(false)} />
              <div className="relative w-72 bg-[#17171c] h-full flex flex-col z-10 border-r border-white/10 animate-slide-in-right">
                <div className="p-4 border-b border-white/10 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <SettingsIcon className="text-[#ffd025]" size={20} />
                    <span className="font-black text-white uppercase text-sm">Menú de Navegación</span>
                  </div>
                  <button type="button" onClick={() => setAdminMobileMenuOpen(false)} className="p-1.5 text-gray-400 hover:text-white">
                    <X size={18} />
                  </button>
                </div>
                <nav className="flex-1 p-3 space-y-1.5 overflow-y-auto custom-admin-scrollbar">
                  {[
                    { id: "products", label: "Productos", icon: Package },
                    { id: "media", label: "Archivo de Imágenes", icon: FolderOpen },
                    { id: "classifications", label: "Clasificaciones", icon: Tag },
                    { id: "contingency", label: "Tienda Contingencia", icon: AlertTriangle },
                    { id: "orders", label: "Pedidos", icon: ClipboardList },
                    { id: "stats", label: "Estadísticas", icon: TrendingUp },
                    { id: "customers", label: "Clientes", icon: Users },
                    { id: "social", label: "Community", icon: Share2 },
                    { id: "settings", label: "Ajustes Generales", icon: SettingsIcon },
                  ]
                    .filter((item) => adminRole === "delivery" ? item.id === "orders" : true)
                    .map((item) => {
                      const Icon = item.icon;
                      const isActive = adminTab === item.id;
                      return (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => {
                            setAdminTab(item.id as any);
                            setAdminMobileMenuOpen(false);
                          }}
                          className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl transition-all text-left ${
                            isActive ? "bg-[#ffd025] text-[#141414] font-bold" : "text-gray-300 hover:bg-white/5"
                          }`}
                        >
                          <Icon size={18} />
                          <span className="text-xs uppercase font-bold">{item.label}</span>
                        </button>
                      );
                    })}
                </nav>
                <div className="p-4 border-t border-white/10 space-y-2">
                  <button
                    type="button"
                    onClick={() => { setView("client"); setAdminMobileMenuOpen(false); }}
                    className="w-full py-2.5 bg-[#ffd025] text-[#141414] font-bold rounded-xl text-xs uppercase flex items-center justify-center gap-2"
                  >
                    <Eye size={15} /> Ver Tienda
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Main Content Workspace */}
          <div className="flex-1 min-w-0 flex flex-col min-h-screen">
            {/* Top Desktop Bar */}
            <header className="hidden md:flex bg-[#17171c]/80 backdrop-blur-md border-b border-white/10 px-6 py-4 items-center justify-between sticky top-0 z-30">
              <div>
                <h1 className="text-lg font-black text-white uppercase tracking-wider flex items-center gap-2">
                  {adminTab === "products" && "📦 Gestión de Productos y Catálogo"}
                  {adminTab === "media" && "📁 Archivo de Imágenes y Galería de Medios"}
                  {adminTab === "classifications" && "🏷️ Clasificaciones, Categorías y Pasillos"}
                  {adminTab === "orders" && "📋 Control de Pedidos y Comandas"}
                  {adminTab === "stats" && "📊 Reportes y Estadísticas de Venta"}
                  {adminTab === "customers" && "👥 Base de Datos de Clientes"}
                  {adminTab === "social" && "📣 Community & Avisos"}
                  {adminTab === "settings" && "⚙️ Ajustes Generales y Personalización"}
                </h1>
                <p className="text-xs text-gray-400">
                  Panel de autoadministración en tiempo real.
                </p>
              </div>
              <div className="flex items-center gap-3">
                {settings.contingencyMode && (
                  <span className="px-3 py-1 bg-red-500/20 text-red-300 border border-red-500/40 rounded-full text-xs font-bold animate-pulse flex items-center gap-1.5">
                    <Ban size={12} /> Modo Contingencia Activo
                  </span>
                )}
                <button
                  type="button"
                  onClick={() => setView("client")}
                  className="flex items-center gap-2 px-4 py-2 bg-[#ffd025] text-[#141414] rounded-xl font-bold hover:bg-[#e5b81a] transition-colors text-xs uppercase shadow-md"
                >
                  <Eye size={15} /> Ver Tienda
                </button>
              </div>
            </header>

            {/* Main Content Padding */}
            <main className="flex-1 p-4 md:p-6 lg:p-8 max-w-7xl w-full mx-auto space-y-6">
              {adminTab === "media" && <MediaAdminPanel />}
              {adminTab === "orders" && <OrdersAdminPanel role={adminRole ?? "full"} />}

            {adminTab === "stats" && <StatsAdminPanel />}
            {adminTab === "customers" && <CustomersAdminPanel />}

            {adminTab === "settings" && (
              <div className="space-y-6 max-w-4xl animate-fade-in">
                {/* Header with quick cloud save and overview */}
                <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-xl shadow-black/60">
                  <div>
                    <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
                      <SettingsIcon size={15} /> Configuración de la Tienda
                    </div>
                    <h2 className="text-xl sm:text-2xl font-black text-white uppercase">Ajustes & Personalización</h2>
                    <p className="text-xs text-gray-400 mt-1">Configura paso a paso o usa los menús plegables para ajustar cada área de tu tienda.</p>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      onClick={() => {
                        const allOpen = Object.values(openSettingsSections).every(Boolean);
                        const nextVal = !allOpen;
                        setOpenSettingsSections({
                          brand: nextVal,
                          ticker: nextVal,
                          banners: nextVal,
                          featured: nextVal,
                          delivery: nextVal,
                          hours: nextVal,
                          contingency: nextVal,
                          contact: nextVal,
                          gallery: nextVal,
                          backup: nextVal,
                        });
                      }}
                      className="px-3.5 py-2.5 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all border border-white/10 flex items-center gap-2"
                    >
                      <Layers size={14} />
                      {Object.values(openSettingsSections).every(Boolean) ? "Plegar Todos" : "Desplegar Todos"}
                    </button>
                    <button
                      type="button"
                      onClick={saveSettings}
                      className="px-5 py-2.5 bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] rounded-xl font-black uppercase text-xs tracking-wider hover:scale-[1.01] active:scale-[0.99] transition-all shadow-lg shadow-[#ffd025]/20 flex items-center gap-2"
                    >
                      <CheckCircle size={15} /> Guardar a la Nube
                    </button>
                  </div>
                </div>

                {/* Step Sub-Tabs (Paso a Paso) */}
                <div className="flex items-center gap-2 overflow-x-auto pb-2 border-b border-white/10 scrollbar-hide">
                  {[
                    { id: "all", label: "Todos los Ajustes", icon: Layers, badge: "10 Plegables" },
                    { id: "brand", label: "1. Marca & Logo", icon: Edit3, badge: settingsDraft.pageTitle || "Tienda" },
                    { id: "ticker", label: "2. Avisos Ticker", icon: Sparkles, badge: `${(settingsDraft.announcements ?? []).length} avisos` },
                    { id: "banners", label: "3. Banners de la Tienda", icon: ImageIcon, badge: `${(settingsDraft.bannerSlides ?? []).length} slides + 2 fijos` },
                    { id: "featured", label: "4. Productos Destacados", icon: Star, badge: `${(settingsDraft.recommendedProductIds ?? []).length} promos / ${(settingsDraft.homeCollectionProductIds ?? []).length} colec.` },
                    { id: "delivery", label: "5. Delivery & Zonas", icon: MapPin, badge: "Tarifas" },
                    { id: "hours", label: "6. Horarios de Pedido", icon: Clock, badge: `${settingsDraft.openTime ?? "11:00"} - ${settingsDraft.closeTime ?? "23:00"}` },
                    { id: "contingency", label: "7. Contingencia", icon: Ban, badge: settingsDraft.contingencyMode ? "🔴 Activo" : "🟢 Normal", highlight: Boolean(settingsDraft.contingencyMode) },
                    { id: "contact", label: "8. Contacto & Redes", icon: Phone, badge: settingsDraft.contactPhone ? "Listo" : "Incompleto" },
                    { id: "gallery", label: "9. Galería 4:5", icon: Camera, badge: `${(settingsDraft.galleryImages ?? []).length || 6} fotos` },
                    { id: "backup", label: "10. Respaldo & Render", icon: ShieldCheck, badge: "Persistencia" },
                  ].map((step) => {
                    const Icon = step.icon;
                    const isActive = settingsSubTab === step.id;
                    return (
                      <button
                        key={step.id}
                        type="button"
                        onClick={() => {
                          setSettingsSubTab(step.id as any);
                          if (step.id !== "all") {
                            setOpenSettingsSections((prev) => ({ ...prev, [step.id]: true }));
                          }
                        }}
                        className={`flex items-center gap-2 px-3.5 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap shrink-0 border ${
                          isActive
                            ? "bg-[#ffd025] text-[#0a0a0f] border-[#ffd025] shadow-lg shadow-[#ffd025]/20"
                            : step.highlight
                            ? "bg-red-500/10 text-red-300 border-red-500/30 hover:bg-red-500/20"
                            : "bg-[#13131f]/60 text-gray-400 border-white/10 hover:text-white hover:border-white/20"
                        }`}
                      >
                        <Icon size={14} className={isActive ? "text-[#0a0a0f]" : "text-[#ffd025]"} />
                        <span>{step.label}</span>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-mono ${isActive ? "bg-black/20 text-[#0a0a0f]" : "bg-white/10 text-gray-400"}`}>
                          {step.badge}
                        </span>
                      </button>
                    );
                  })}
                </div>

                <div className="space-y-4">
                  {/* SECCIÓN 1: Identidad & Branding */}
                  {(settingsSubTab === "all" || settingsSubTab === "brand") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("brand")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center shrink-0">
                            <Edit3 size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 1: Identidad & Marca</span>
                              <span className="text-[10px] bg-amber-500/15 text-amber-300 border border-amber-500/30 px-2 py-0.5 rounded-full font-bold">
                                {settingsDraft.pageTitle || "Configurar"}
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Título del navegador, logo principal y favicon.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.brand ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.brand && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-4 animate-fade-in">
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4">
                            <div>
                              <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                                Título de la página (Pestaña)
                              </label>
                              <input
                                type="text"
                                value={settingsDraft.pageTitle || ""}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, pageTitle: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] focus:ring-2 focus:ring-[#ffd025]/20 text-sm placeholder-gray-500"
                                placeholder="Ej: Urban Bite - Delivery"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                                Favicon (Ícono cuadrado de pestaña)
                              </label>
                              <div className="flex gap-2">
                                <input
                                  type="text"
                                  value={settingsDraft.favicon || ""}
                                  onChange={(e) =>
                                    setSettingsDraft({ ...settingsDraft, favicon: e.target.value })
                                  }
                                  onBlur={(e) =>
                                    resolveImageUrl(e.target.value, (r) =>
                                      setSettingsDraft((p) => ({ ...p, favicon: r })),
                                    )
                                  }
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm placeholder-gray-500"
                                  placeholder="URL o sube una imagen"
                                />
                                <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                                  <Upload size={18} />
                                  <input
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={(e) =>
                                      handleImageUpload(e, (url) =>
                                        setSettingsDraft({ ...settingsDraft, favicon: url }),
                                      )
                                    }
                                  />
                                </label>
                                {(settingsDraft.favicon || "").includes("/storage/objects/") && (
                                  <button
                                    type="button"
                                    title="Eliminar imagen"
                                    onClick={() =>
                                      handleDeleteStorageImage(settingsDraft.favicon || "", () =>
                                        setSettingsDraft((p) => ({ ...p, favicon: "" })),
                                      )
                                    }
                                    className="bg-red-500/10 text-red-400 px-3.5 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30"
                                  >
                                    <Trash2 size={16} />
                                  </button>
                                )}
                              </div>
                            </div>
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                              Logo Principal de la Tienda
                            </label>
                            <div className="flex gap-3 items-center">
                              {settingsDraft.logo && (
                                <img
                                  src={settingsDraft.logo}
                                  alt="Logo"
                                  className="w-12 h-12 rounded-xl object-contain bg-black border border-white/10 p-1 shrink-0"
                                />
                              )}
                              <input
                                type="text"
                                value={settingsDraft.logo || ""}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, logo: e.target.value })
                                }
                                onBlur={(e) =>
                                  resolveImageUrl(e.target.value, (r) =>
                                    setSettingsDraft((p) => ({ ...p, logo: r })),
                                  )
                                }
                                className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm placeholder-gray-500"
                                placeholder="URL del logo de tu tienda"
                              />
                              <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 py-3 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30 shrink-0">
                                <Upload size={18} />
                                <input
                                  type="file"
                                  accept="image/*"
                                  className="hidden"
                                  onChange={(e) =>
                                    handleImageUpload(e, (url) =>
                                      setSettingsDraft({ ...settingsDraft, logo: url }),
                                    )
                                  }
                                />
                              </label>
                              {(settingsDraft.logo || "").includes("/storage/objects/") && (
                                <button
                                  type="button"
                                  title="Eliminar logo"
                                  onClick={() =>
                                    handleDeleteStorageImage(settingsDraft.logo || "", () =>
                                      setSettingsDraft((p) => ({ ...p, logo: "" })),
                                    )
                                  }
                                  className="bg-red-500/10 text-red-400 px-3.5 py-3 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30 shrink-0"
                                >
                                  <Trash2 size={16} />
                                </button>
                              )}
                            </div>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("ticker");
                                setOpenSettingsSections((p) => ({ ...p, ticker: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Avisos Ticker <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 2: Barra Ticker de Avisos */}
                  {(settingsSubTab === "all" || settingsSubTab === "ticker") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("ticker")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-purple-500/10 border border-purple-500/20 text-purple-400 flex items-center justify-center shrink-0">
                            <Sparkles size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 2: Barra de Avisos (Ticker Animado)</span>
                              <span className="text-[10px] bg-purple-500/15 text-purple-300 border border-purple-500/30 px-2 py-0.5 rounded-full font-bold">
                                {(settingsDraft.announcements ?? []).length} activos
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Mensajes deslizantes en el tope de la tienda. Con optimizador IA.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.ticker ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.ticker && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-5 animate-fade-in">
                          {/* 1. Anuncio Breve Superior (Barra Degradada sobre el Encabezado) */}
                          <div className="bg-[#181826] p-4 rounded-2xl border border-white/10 space-y-2 mt-4">
                            <div className="flex items-center justify-between">
                              <label className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                Anuncio Breve Superior (Barra sobre el Encabezado)
                              </label>
                              <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 font-bold border border-amber-500/30">
                                Visible siempre
                              </span>
                            </div>
                            <p className="text-[11px] text-gray-400">
                              Texto que aparece en la barra superior ultra delgada con degradado naranja sobre el logo y buscador.
                            </p>
                            <input
                              type="text"
                              value={settingsDraft.topAnnouncementText ?? ""}
                              onChange={(e) => setSettingsDraft({ ...settingsDraft, topAnnouncementText: e.target.value })}
                              placeholder="Ej: PIDE ANTES DE LAS 8:00 AM Y RECIBE EL MISMO DÍA (VER COMUNAS)"
                              className="w-full bg-[#12121d] border border-white/10 rounded-xl p-3 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                            />
                          </div>

                          {/* 2. Mensajes Rotativos del Ticker */}
                          <div className="space-y-3">
                            <div className="flex items-center justify-between">
                              <span className="text-xs font-black text-purple-300 uppercase tracking-wider block">
                                Mensajes Rotativos de Avisos
                              </span>
                              <span className="text-[10px] text-gray-400">Rotan en la tienda</span>
                            </div>
                            <p className="text-xs text-gray-400">
                              Avisos adicionales que rotan periódicamente. Puedes potenciar cada aviso con IA para mayor engagement.
                            </p>
                            <div className="space-y-2">
                            {((settingsDraft.announcements ?? []) as string[]).map((txt, idx) => (
                              <div key={idx} className="flex gap-2 items-center">
                                <input
                                  type="text"
                                  value={txt}
                                  onChange={(e) => {
                                    const arr = [...((settingsDraft.announcements ?? []) as string[])];
                                    arr[idx] = e.target.value;
                                    setSettingsDraft({ ...settingsDraft, announcements: arr });
                                  }}
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3 text-white text-xs focus:border-purple-400 focus:outline-none"
                                  placeholder="Ej: 🔥 Despacho gratis sobre $20.000 #promo"
                                />
                                <button
                                  type="button"
                                  onClick={async () => {
                                    try {
                                      showToast("Mejorando con IA...");
                                      const resp = await fetch(`${import.meta.env.BASE_URL}api/ai/improve-announcement`, {
                                        method: "POST",
                                        headers: { "Content-Type": "application/json" },
                                        body: JSON.stringify({ text: txt }),
                                      });
                                      const data = await resp.json();
                                      if (data.text) {
                                        const arr = [...((settingsDraft.announcements ?? []) as string[])];
                                        arr[idx] = data.text;
                                        setSettingsDraft({ ...settingsDraft, announcements: arr });
                                        showToast("Aviso mejorado con IA ✨");
                                      }
                                    } catch {
                                      showToast("No se pudo mejorar el aviso");
                                    }
                                  }}
                                  className="p-3 bg-purple-500/15 text-purple-300 hover:bg-purple-500/30 rounded-xl transition-colors border border-purple-500/30 shrink-0"
                                  title="Mejorar redacción con IA"
                                >
                                  <Sparkles size={16} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    const arr = [...((settingsDraft.announcements ?? []) as string[])];
                                    arr.splice(idx, 1);
                                    setSettingsDraft({ ...settingsDraft, announcements: arr });
                                  }}
                                  className="p-3 bg-red-500/10 text-red-400 hover:bg-red-500/20 rounded-xl transition-colors border border-red-500/30 shrink-0"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </div>
                            ))}
                            <button
                              type="button"
                              onClick={() => {
                                const arr = [...((settingsDraft.announcements ?? []) as string[])];
                                arr.push("");
                                setSettingsDraft({ ...settingsDraft, announcements: arr });
                              }}
                              className="w-full py-3 bg-purple-500/10 text-purple-300 hover:bg-purple-500/20 rounded-xl text-xs font-bold uppercase flex items-center justify-center gap-2 border border-purple-500/20 transition-colors"
                            >
                              <Plus size={14} /> Agregar Nuevo Aviso
                            </button>
                            </div>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("banners");
                                setOpenSettingsSections((p) => ({ ...p, banners: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Carrusel Banners <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 3: Gestión de Banners de la Tienda */}
                  {(settingsSubTab === "all" || settingsSubTab === "banners") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("banners")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-cyan-500/10 border border-cyan-500/20 text-cyan-400 flex items-center justify-center shrink-0">
                            <ImageIcon size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 3: Banners de la Tienda (Portada, Promos y Pasillos)</span>
                              <span className="text-[10px] bg-cyan-500/15 text-cyan-300 border border-cyan-500/30 px-2 py-0.5 rounded-full font-bold">
                                {(settingsDraft.bannerSlides ?? []).length} slides + 2 banners
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Modifica el banner de promociones, el banner de pasillos y el carrusel de portada.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.banners ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.banners && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-6 animate-fade-in">
                          {/* Banner 1: Banner de Promociones (#NUESTROSRECOMENDADOS) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-3 mt-4">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div>
                                <span className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                  1. Banner Sección Promociones (#NUESTROSRECOMENDADOS)
                                </span>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Imagen panorámica horizontal que encabeza las promociones y recomendaciones del Tío Fellas.
                                </p>
                              </div>
                              <span className="text-[10px] px-2 py-0.5 rounded-full bg-red-500/15 text-red-300 border border-red-500/30 font-bold uppercase">
                                Promociones
                              </span>
                            </div>

                            <div className="w-full h-20 sm:h-28 md:h-36 overflow-hidden rounded-xl border border-white/10 bg-black relative select-none">
                              <img
                                src={settingsDraft.promoBannerImage || "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=1600&auto=format&fit=crop&q=80"}
                                alt="Banner Promociones"
                                className="w-full h-full object-cover"
                              />
                            </div>

                            <div className="flex gap-2 items-center">
                              <input
                                type="text"
                                value={settingsDraft.promoBannerImage || ""}
                                onChange={(e) => setSettingsDraft({ ...settingsDraft, promoBannerImage: e.target.value })}
                                onBlur={(e) =>
                                  resolveImageUrl(e.target.value, (r) =>
                                    setSettingsDraft((p) => ({ ...p, promoBannerImage: r }))
                                  )
                                }
                                className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                placeholder="URL imagen del banner de promociones (o presiona Subir)"
                              />
                              <label
                                className="bg-[#ffd025]/10 text-[#ffd025] px-3.5 py-2.5 rounded-xl flex items-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/20 shrink-0 font-bold text-xs gap-1.5"
                                title="Subir imagen desde tu dispositivo"
                              >
                                <Upload size={14} />
                                <span className="hidden sm:inline">Subir</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  className="hidden"
                                  onChange={(e) =>
                                    handleImageUpload(e, (url) =>
                                      setSettingsDraft({ ...settingsDraft, promoBannerImage: url })
                                    )
                                  }
                                />
                              </label>
                              {settingsDraft.promoBannerImage && (
                                <button
                                  type="button"
                                  onClick={() => setSettingsDraft({ ...settingsDraft, promoBannerImage: "" })}
                                  className="p-2.5 bg-red-500/10 text-red-400 hover:bg-red-500/20 rounded-xl transition-colors border border-red-500/20 text-xs shrink-0"
                                  title="Restaurar por defecto"
                                >
                                  <Trash2 size={15} />
                                </button>
                              )}
                            </div>
                          </div>

                          {/* Banner 2: Banner de Pasillos y Colecciones (#NUESTRASCOLECCIONES) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-3">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div>
                                <span className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                  2. Banner Sección Colecciones & Pasillos (#NUESTRASCOLECCIONES)
                                </span>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Imagen panorámica horizontal que se muestra sobre los pasillos en la portada y en la pestaña de productos.
                                </p>
                              </div>
                              <span className="text-[10px] px-2 py-0.5 rounded-full bg-cyan-500/15 text-cyan-300 border border-cyan-500/30 font-bold uppercase">
                                Pasillos
                              </span>
                            </div>

                            <div className="w-full h-20 sm:h-28 md:h-36 overflow-hidden rounded-xl border border-white/10 bg-black relative select-none">
                              <img
                                src={settingsDraft.aislesBannerImage || "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=1600&auto=format&fit=crop&q=80"}
                                alt="Banner Pasillos"
                                className="w-full h-full object-cover"
                              />
                            </div>

                            <div className="flex gap-2 items-center">
                              <input
                                type="text"
                                value={settingsDraft.aislesBannerImage || ""}
                                onChange={(e) => setSettingsDraft({ ...settingsDraft, aislesBannerImage: e.target.value })}
                                onBlur={(e) =>
                                  resolveImageUrl(e.target.value, (r) =>
                                    setSettingsDraft((p) => ({ ...p, aislesBannerImage: r }))
                                  )
                                }
                                className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                placeholder="URL imagen del banner de colecciones y pasillos"
                              />
                              <label
                                className="bg-[#ffd025]/10 text-[#ffd025] px-3.5 py-2.5 rounded-xl flex items-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/20 shrink-0 font-bold text-xs gap-1.5"
                                title="Subir imagen desde tu dispositivo"
                              >
                                <Upload size={14} />
                                <span className="hidden sm:inline">Subir</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  className="hidden"
                                  onChange={(e) =>
                                    handleImageUpload(e, (url) =>
                                      setSettingsDraft({ ...settingsDraft, aislesBannerImage: url })
                                    )
                                  }
                                />
                              </label>
                              {settingsDraft.aislesBannerImage && (
                                <button
                                  type="button"
                                  onClick={() => setSettingsDraft({ ...settingsDraft, aislesBannerImage: "" })}
                                  className="p-2.5 bg-red-500/10 text-red-400 hover:bg-red-500/20 rounded-xl transition-colors border border-red-500/20 text-xs shrink-0"
                                  title="Restaurar por defecto"
                                >
                                  <Trash2 size={15} />
                                </button>
                              )}
                            </div>
                          </div>

                          {/* Banner 3: Carrusel de Banners de Portada (Slides de Inicio) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-4">
                            <div className="flex items-center justify-between">
                              <div>
                                <span className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                  3. Carrusel de Banners de Portada (Slides)
                                </span>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Slides que rotan automáticamente en la parte superior de la página principal.
                                </p>
                              </div>
                              <button
                                type="button"
                                onClick={() => {
                                  const newSlide: BannerSlide = { image: "", title: "", description: "" };
                                  setSettingsDraft({
                                    ...settingsDraft,
                                    bannerSlides: [...(settingsDraft.bannerSlides ?? []), newSlide],
                                  });
                                }}
                                className="flex items-center gap-1.5 text-xs bg-[#ffd025]/10 text-[#ffd025] border border-[#ffd025]/30 rounded-xl px-3 py-1.5 hover:bg-[#ffd025]/20 transition-colors font-bold"
                              >
                                <Plus size={14} /> Añadir slide
                              </button>
                            </div>

                            {(settingsDraft.bannerSlides ?? []).length === 0 && (
                              <div className="text-center py-6 border border-dashed border-white/10 rounded-2xl text-gray-500 text-xs">
                                Sin slides dinámicos configurados.
                              </div>
                            )}

                            <div className="space-y-3">
                              {(settingsDraft.bannerSlides ?? []).map((slide, idx) => (
                                <div
                                  key={idx}
                                  className="rounded-2xl border border-white/10 bg-[#141420] p-3.5 sm:p-4 space-y-3"
                                >
                                  <div className="flex items-center justify-between">
                                    <span className="text-[10px] font-black text-[#ffd025] uppercase tracking-wider">
                                      Slide {idx + 1}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const slides = [...(settingsDraft.bannerSlides ?? [])];
                                        slides.splice(idx, 1);
                                        setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                      }}
                                      className="text-red-400 hover:text-red-300 text-xs font-bold"
                                    >
                                      Eliminar
                                    </button>
                                  </div>
                                  <div className="flex gap-2 items-center">
                                    {slide.image && (
                                      <img src={slide.image} alt="" className="w-16 h-10 object-cover rounded-lg bg-black border border-white/10 shrink-0" />
                                    )}
                                    <input
                                      type="text"
                                      value={slide.image}
                                      onChange={(e) => {
                                        const slides = [...(settingsDraft.bannerSlides ?? [])];
                                        slides[idx] = { ...slides[idx]!, image: e.target.value };
                                        setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                      }}
                                      onBlur={(e) =>
                                        resolveImageUrl(e.target.value, (r) => {
                                          const slides = [...(settingsDraft.bannerSlides ?? [])];
                                          slides[idx] = { ...slides[idx]!, image: r };
                                          setSettingsDraft((p) => ({ ...p, bannerSlides: slides }));
                                        })
                                      }
                                      className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                      placeholder="URL imagen del slide"
                                    />
                                    <label className="bg-[#ffd025]/10 text-[#ffd025] px-3.5 py-2.5 rounded-xl flex items-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/20 shrink-0">
                                      <Upload size={15} />
                                      <input
                                        type="file"
                                        accept="image/*"
                                        className="hidden"
                                        onChange={(e) =>
                                          handleImageUpload(e, (url) => {
                                            const slides = [...(settingsDraft.bannerSlides ?? [])];
                                            slides[idx] = { ...slides[idx]!, image: url };
                                            setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                          })
                                        }
                                      />
                                    </label>
                                  </div>
                                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                    <input
                                      type="text"
                                      value={slide.title}
                                      onChange={(e) => {
                                        const slides = [...(settingsDraft.bannerSlides ?? [])];
                                        slides[idx] = { ...slides[idx]!, title: e.target.value };
                                        setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                      }}
                                      className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                      placeholder="Título (opcional)"
                                    />
                                    <input
                                      type="text"
                                      value={slide.description}
                                      onChange={(e) => {
                                        const slides = [...(settingsDraft.bannerSlides ?? [])];
                                        slides[idx] = { ...slides[idx]!, description: e.target.value };
                                        setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                      }}
                                      className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                      placeholder="Descripción corta (opcional)"
                                    />
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("featured");
                                setOpenSettingsSections((p) => ({ ...p, featured: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Productos Destacados <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 4: Productos Destacados (#NUESTROSRECOMENDADOS y #NUESTRASCOLECCIONES) */}
                  {(settingsSubTab === "all" || settingsSubTab === "featured") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("featured")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center shrink-0">
                            <Star size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 4: Selección de Productos Destacados</span>
                              <span className="text-[10px] bg-amber-500/15 text-amber-300 border border-amber-500/30 px-2 py-0.5 rounded-full font-bold">
                                {(settingsDraft.recommendedProductIds ?? []).length} promos / {(settingsDraft.homeCollectionProductIds ?? []).length} colec.
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">
                              Elige manualmente los productos de Promociones del Tío Fellas (hasta 4 en PC / 3 en celular) y los 6 productos iniciales de Colecciones antes de abrir el catálogo completo.
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.featured ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.featured && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-8 animate-fade-in">
                          {/* BLOQUE A: #NUESTROSRECOMENDADOS (Promos del Tío Fellas) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-4 mt-4">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                    A. Promociones del Tío Fellas (#NUESTROSRECOMENDADOS)
                                  </span>
                                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-red-500/20 text-red-300 font-bold border border-red-500/30">
                                    {(settingsDraft.recommendedProductIds ?? []).length} seleccionados
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Estos productos aparecen en el bloque destacado con la etiqueta roja PROMO bajo el banner de ofertas.
                                </p>
                              </div>
                              <div className="flex items-center gap-2 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const candidates = products.filter((p) => !p.hidden && (p.oferta || p.bestseller)).map((p) => p.id);
                                    const fallback = products.filter((p) => !p.hidden).map((p) => p.id);
                                    const picked = Array.from(new Set([...candidates, ...fallback])).slice(0, 4);
                                    setSettingsDraft((p) => ({ ...p, recommendedProductIds: picked }));
                                    showToast("Cargadas 4 ofertas para recomendaciones");
                                  }}
                                  className="px-3 py-1.5 bg-[#ffd025]/10 hover:bg-[#ffd025]/20 text-[#ffd025] border border-[#ffd025]/30 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Cargar Ofertas
                                </button>
                                {(settingsDraft.recommendedProductIds ?? []).length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setSettingsDraft((p) => ({ ...p, recommendedProductIds: [] }));
                                      showToast("Selección de promociones limpiada");
                                    }}
                                    className="px-3 py-1.5 bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white rounded-xl text-[10px] font-bold uppercase transition-colors border border-white/10"
                                  >
                                    Limpiar
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Lista de productos actualmente seleccionados para Promos */}
                            <div>
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block mb-2">
                                Productos actualmente en #NUESTROSRECOMENDADOS (En orden de aparición):
                              </span>
                              {(settingsDraft.recommendedProductIds ?? []).length === 0 ? (
                                <div className="text-center py-5 border border-dashed border-white/10 rounded-xl text-gray-400 text-xs">
                                  Sin productos seleccionados manualmente. (El sistema mostrará automáticamente las ofertas del catálogo).
                                </div>
                              ) : (
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                                  {(settingsDraft.recommendedProductIds ?? []).map((id, idx) => {
                                    const prod = products.find((p) => p.id === id);
                                    if (!prod) return null;
                                    return (
                                      <div
                                        key={id}
                                        className="bg-[#12121d] border border-white/10 rounded-xl p-2.5 flex items-center justify-between gap-2 shadow"
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          <span className="text-[10px] font-black text-[#ffd025] px-1.5 py-0.5 bg-black/40 rounded border border-white/10">
                                            #{idx + 1}
                                          </span>
                                          {prod.image ? (
                                            <img
                                              src={prod.image}
                                              alt={prod.name}
                                              className="w-10 h-10 object-cover rounded-lg bg-black shrink-0 border border-white/10"
                                            />
                                          ) : (
                                            <div className="w-10 h-10 bg-white/5 rounded-lg flex items-center justify-center shrink-0">
                                              <Package size={16} className="text-gray-500" />
                                            </div>
                                          )}
                                          <div className="min-w-0">
                                            <p className="text-xs font-bold text-white truncate">{prod.name}</p>
                                            <p className="text-[10px] text-[#ffd025] font-black">
                                              ${Number(prod.price || 0).toLocaleString("es-CL")}
                                            </p>
                                          </div>
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          {idx > 0 && (
                                            <button
                                              type="button"
                                              onClick={() => {
                                                const list = [...(settingsDraft.recommendedProductIds ?? [])];
                                                const tmp = list[idx];
                                                list[idx] = list[idx - 1]!;
                                                list[idx - 1] = tmp!;
                                                setSettingsDraft((p) => ({ ...p, recommendedProductIds: list }));
                                              }}
                                              className="p-1 text-gray-400 hover:text-white hover:bg-white/10 rounded"
                                              title="Mover a la izquierda"
                                            >
                                              <ArrowLeft size={13} />
                                            </button>
                                          )}
                                          {idx < (settingsDraft.recommendedProductIds ?? []).length - 1 && (
                                            <button
                                              type="button"
                                              onClick={() => {
                                                const list = [...(settingsDraft.recommendedProductIds ?? [])];
                                                const tmp = list[idx];
                                                list[idx] = list[idx + 1]!;
                                                list[idx + 1] = tmp!;
                                                setSettingsDraft((p) => ({ ...p, recommendedProductIds: list }));
                                              }}
                                              className="p-1 text-gray-400 hover:text-white hover:bg-white/10 rounded"
                                              title="Mover a la derecha"
                                            >
                                              <ChevronRight size={13} />
                                            </button>
                                          )}
                                          <button
                                            type="button"
                                            onClick={() => {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                recommendedProductIds: (p.recommendedProductIds ?? []).filter((x) => x !== id),
                                              }));
                                            }}
                                            className="p-1 text-red-400 hover:text-red-300 hover:bg-red-500/10 rounded"
                                            title="Quitar de recomendaciones"
                                          >
                                            <X size={14} />
                                          </button>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>

                            {/* Buscador de catálogo para añadir a Promos */}
                            <div className="pt-2 border-t border-white/5 space-y-2">
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block">
                                Buscar y añadir productos del catálogo a Promociones:
                              </span>
                              <input
                                type="text"
                                value={featuredRecoSearch}
                                onChange={(e) => setFeaturedRecoSearch(e.target.value)}
                                placeholder="Filtrar por nombre, pasillo, pack, cerveza, pisco..."
                                className="w-full bg-[#12121d] border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:border-[#ffd025]"
                              />
                              <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1">
                                {products
                                  .filter((p) => !p.hidden)
                                  .filter((p) => {
                                    if (!featuredRecoSearch.trim()) return true;
                                    const q = featuredRecoSearch.toLowerCase();
                                    return (
                                      p.name.toLowerCase().includes(q) ||
                                      (p.category && p.category.toLowerCase().includes(q)) ||
                                      (p.aisle && p.aisle.toLowerCase().includes(q))
                                    );
                                  })
                                  .slice(0, 30)
                                  .map((product) => {
                                    const isSelected = (settingsDraft.recommendedProductIds ?? []).includes(product.id);
                                    return (
                                      <div
                                        key={product.id}
                                        className={`flex items-center justify-between p-2 rounded-xl border transition-all ${
                                          isSelected
                                            ? "bg-[#ffd025]/10 border-[#ffd025]/30 text-white"
                                            : "bg-[#141420] border-white/5 text-gray-300 hover:border-white/15"
                                        }`}
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          {product.image ? (
                                            <img
                                              src={product.image}
                                              alt={product.name}
                                              className="w-8 h-8 object-cover rounded-lg bg-black shrink-0 border border-white/10"
                                            />
                                          ) : (
                                            <div className="w-8 h-8 bg-white/5 rounded-lg flex items-center justify-center shrink-0">
                                              <Package size={14} className="text-gray-500" />
                                            </div>
                                          )}
                                          <div className="min-w-0">
                                            <p className="text-xs font-bold text-white truncate">{product.name}</p>
                                            <p className="text-[10px] text-gray-400">
                                              ${Number(product.price || 0).toLocaleString("es-CL")} · {product.aisle || product.category}
                                            </p>
                                          </div>
                                        </div>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (isSelected) {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                recommendedProductIds: (p.recommendedProductIds ?? []).filter((x) => x !== product.id),
                                              }));
                                            } else {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                recommendedProductIds: [...(p.recommendedProductIds ?? []), product.id],
                                              }));
                                              showToast(`"${product.name}" añadido a Promociones`);
                                            }
                                          }}
                                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                                            isSelected
                                              ? "bg-red-500/20 text-red-300 border border-red-500/30 hover:bg-red-500/30"
                                              : "bg-[#ffd025] text-black font-black hover:bg-[#e5b81a]"
                                          }`}
                                        >
                                          {isSelected ? "Quitar" : "+ Añadir a Promos"}
                                        </button>
                                      </div>
                                    );
                                  })}
                              </div>
                            </div>
                          </div>

                          {/* BLOQUE B: #NUESTRASCOLECCIONES (Productos destacados para Celular y PC) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-4">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black text-[#ffd025] uppercase tracking-wider block">
                                    B. Productos Destacados en #NUESTRASCOLECCIONES (Vista Celular y PC)
                                  </span>
                                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-cyan-500/20 text-cyan-300 font-bold border border-cyan-500/30">
                                    {(settingsDraft.homeCollectionProductIds ?? []).length} / 6 productos
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Estos son exactamente los 6 productos destacados que verá el cliente en su celular y en computador antes de hacer clic en "Ver más" / "Ver catálogo completo".
                                </p>
                              </div>
                              <div className="flex items-center gap-2 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const first6 = products.filter((p) => !p.hidden).slice(0, 6).map((p) => p.id);
                                    setSettingsDraft((p) => ({ ...p, homeCollectionProductIds: first6 }));
                                    showToast("Cargados los primeros 6 del catálogo");
                                  }}
                                  className="px-3 py-1.5 bg-[#ffd025]/10 hover:bg-[#ffd025]/20 text-[#ffd025] border border-[#ffd025]/30 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Cargar Primeros 6
                                </button>
                                {(settingsDraft.homeCollectionProductIds ?? []).length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setSettingsDraft((p) => ({ ...p, homeCollectionProductIds: [] }));
                                      showToast("Selección de colecciones limpiada");
                                    }}
                                    className="px-3 py-1.5 bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white rounded-xl text-[10px] font-bold uppercase transition-colors border border-white/10"
                                  >
                                    Limpiar
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Lista visual de las 6 posiciones fijas */}
                            <div>
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block mb-2">
                                Vista previa de los 6 productos iniciales para móvil:
                              </span>
                              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2">
                                {Array.from({ length: 6 }).map((_, slotIdx) => {
                                  const prodId = (settingsDraft.homeCollectionProductIds ?? [])[slotIdx];
                                  const prod = prodId ? products.find((p) => p.id === prodId) : null;

                                  if (!prod) {
                                    return (
                                      <div
                                        key={slotIdx}
                                        className="h-28 rounded-xl border border-dashed border-white/10 bg-black/20 p-2 flex flex-col items-center justify-center text-center"
                                      >
                                        <span className="text-[10px] font-black text-gray-500 mb-1">Posición #{slotIdx + 1}</span>
                                        <span className="text-[10px] text-gray-500 italic">Espacio disponible</span>
                                      </div>
                                    );
                                  }

                                  return (
                                    <div
                                      key={prod.id}
                                      className="rounded-xl border border-white/10 bg-[#12121d] p-2 flex flex-col justify-between shadow relative group"
                                    >
                                      <div>
                                        <div className="relative aspect-square w-full rounded-lg overflow-hidden bg-black mb-1.5">
                                          {prod.image ? (
                                            <img src={prod.image} alt={prod.name} className="w-full h-full object-cover" />
                                          ) : (
                                            <div className="w-full h-full flex items-center justify-center bg-white/5">
                                              <Package size={16} className="text-gray-500" />
                                            </div>
                                          )}
                                          <span className="absolute top-1 left-1 px-1.5 py-0.5 bg-[#ffd025] text-black text-[9px] font-black rounded">
                                            #{slotIdx + 1}
                                          </span>
                                        </div>
                                        <p className="text-[11px] font-bold text-white line-clamp-1 leading-tight">{prod.name}</p>
                                        <p className="text-[10px] text-[#ffd025] font-black">
                                          ${Number(prod.price || 0).toLocaleString("es-CL")}
                                        </p>
                                      </div>
                                      <div className="flex items-center justify-between pt-1.5 mt-1 border-t border-white/10">
                                        <div className="flex items-center gap-1">
                                          {slotIdx > 0 && (
                                            <button
                                              type="button"
                                              onClick={() => {
                                                const list = [...(settingsDraft.homeCollectionProductIds ?? [])];
                                                const tmp = list[slotIdx];
                                                list[slotIdx] = list[slotIdx - 1]!;
                                                list[slotIdx - 1] = tmp!;
                                                setSettingsDraft((p) => ({ ...p, homeCollectionProductIds: list }));
                                              }}
                                              className="p-0.5 text-gray-400 hover:text-white"
                                              title="Mover a posición anterior"
                                            >
                                              <ArrowLeft size={12} />
                                            </button>
                                          )}
                                          {slotIdx < (settingsDraft.homeCollectionProductIds ?? []).length - 1 && (
                                            <button
                                              type="button"
                                              onClick={() => {
                                                const list = [...(settingsDraft.homeCollectionProductIds ?? [])];
                                                const tmp = list[slotIdx];
                                                list[slotIdx] = list[slotIdx + 1]!;
                                                list[slotIdx + 1] = tmp!;
                                                setSettingsDraft((p) => ({ ...p, homeCollectionProductIds: list }));
                                              }}
                                              className="p-0.5 text-gray-400 hover:text-white"
                                              title="Mover a posición siguiente"
                                            >
                                              <ChevronRight size={12} />
                                            </button>
                                          )}
                                        </div>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setSettingsDraft((p) => ({
                                              ...p,
                                              homeCollectionProductIds: (p.homeCollectionProductIds ?? []).filter((x) => x !== prod.id),
                                            }));
                                          }}
                                          className="p-0.5 text-red-400 hover:text-red-300"
                                          title="Quitar de la colección inicial"
                                        >
                                          <X size={13} />
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>

                            {/* Buscador de catálogo para añadir a Colecciones */}
                            <div className="pt-2 border-t border-white/5 space-y-2">
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block">
                                Buscar y añadir productos a los 6 iniciales:
                              </span>
                              <input
                                type="text"
                                value={featuredCollecSearch}
                                onChange={(e) => setFeaturedCollecSearch(e.target.value)}
                                placeholder="Filtrar por nombre, categoría, pasillo..."
                                className="w-full bg-[#12121d] border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:border-[#ffd025]"
                              />
                              <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1">
                                {products
                                  .filter((p) => !p.hidden)
                                  .filter((p) => {
                                    if (!featuredCollecSearch.trim()) return true;
                                    const q = featuredCollecSearch.toLowerCase();
                                    return (
                                      p.name.toLowerCase().includes(q) ||
                                      (p.category && p.category.toLowerCase().includes(q)) ||
                                      (p.aisle && p.aisle.toLowerCase().includes(q))
                                    );
                                  })
                                  .slice(0, 30)
                                  .map((product) => {
                                    const isSelected = (settingsDraft.homeCollectionProductIds ?? []).includes(product.id);
                                    const isFull = (settingsDraft.homeCollectionProductIds ?? []).length >= 6;
                                    return (
                                      <div
                                        key={product.id}
                                        className={`flex items-center justify-between p-2 rounded-xl border transition-all ${
                                          isSelected
                                            ? "bg-cyan-500/10 border-cyan-500/30 text-white"
                                            : "bg-[#141420] border-white/5 text-gray-300 hover:border-white/15"
                                        }`}
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          {product.image ? (
                                            <img
                                              src={product.image}
                                              alt={product.name}
                                              className="w-8 h-8 object-cover rounded-lg bg-black shrink-0 border border-white/10"
                                            />
                                          ) : (
                                            <div className="w-8 h-8 bg-white/5 rounded-lg flex items-center justify-center shrink-0">
                                              <Package size={14} className="text-gray-500" />
                                            </div>
                                          )}
                                          <div className="min-w-0">
                                            <p className="text-xs font-bold text-white truncate">{product.name}</p>
                                            <p className="text-[10px] text-gray-400">
                                              ${Number(product.price || 0).toLocaleString("es-CL")} · {product.aisle || product.category}
                                            </p>
                                          </div>
                                        </div>
                                        <button
                                          type="button"
                                          disabled={!isSelected && isFull}
                                          onClick={() => {
                                            if (isSelected) {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                homeCollectionProductIds: (p.homeCollectionProductIds ?? []).filter((x) => x !== product.id),
                                              }));
                                            } else {
                                              if (isFull) {
                                                showToast("Ya alcanzaste el máximo de 6 productos");
                                                return;
                                              }
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                homeCollectionProductIds: [...(p.homeCollectionProductIds ?? []), product.id],
                                              }));
                                              showToast(`"${product.name}" agregado a los 6 iniciales`);
                                            }
                                          }}
                                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                                            isSelected
                                              ? "bg-red-500/20 text-red-300 border border-red-500/30 hover:bg-red-500/30"
                                              : isFull
                                              ? "bg-white/5 text-gray-500 cursor-not-allowed border border-white/5"
                                              : "bg-[#ffd025] text-black font-black hover:bg-[#e5b81a]"
                                          }`}
                                        >
                                          {isSelected ? "Quitar" : isFull ? "Límite 6 alcanzado" : "+ Añadir a Colección"}
                                        </button>
                                      </div>
                                    );
                                  })}
                              </div>
                            </div>
                          </div>

                          {/* BLOQUE C: Pestaña Exclusiva Oportunidades ⏰ (Banner + Selección de Productos) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-4">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black text-amber-400 uppercase tracking-wider block">
                                    C. Pestaña Exclusiva "OPORTUNIDADES" ⏰ (Banner + Selección de Productos)
                                  </span>
                                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 font-bold border border-amber-500/30">
                                    {(settingsDraft.opportunitiesProductIds ?? []).length} seleccionados
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Configura el banner panorámico y elige la lista de productos (5, 10, 20 o más) que se mostrarán al presionar "OPORTUNIDADES".
                                </p>
                              </div>
                              <div className="flex items-center gap-2 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const offers = products.filter((p) => !p.hidden && (p.oferta || p.bestseller)).map((p) => p.id);
                                    const fallback = products.filter((p) => !p.hidden).map((p) => p.id);
                                    const picked = Array.from(new Set([...offers, ...fallback])).slice(0, 10);
                                    setSettingsDraft((p) => ({ ...p, opportunitiesProductIds: picked }));
                                    showToast("Cargados 10 productos para Oportunidades");
                                  }}
                                  className="px-3 py-1.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Cargar 10
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    const offers = products.filter((p) => !p.hidden && (p.oferta || p.bestseller)).map((p) => p.id);
                                    const fallback = products.filter((p) => !p.hidden).map((p) => p.id);
                                    const picked = Array.from(new Set([...offers, ...fallback])).slice(0, 20);
                                    setSettingsDraft((p) => ({ ...p, opportunitiesProductIds: picked }));
                                    showToast("Cargados 20 productos para Oportunidades");
                                  }}
                                  className="px-3 py-1.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Cargar 20
                                </button>
                                {(settingsDraft.opportunitiesProductIds ?? []).length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setSettingsDraft((p) => ({ ...p, opportunitiesProductIds: [] }));
                                      showToast("Lista de Oportunidades limpiada");
                                    }}
                                    className="px-3 py-1.5 bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white rounded-xl text-[10px] font-bold uppercase transition-colors border border-white/10"
                                  >
                                    Limpiar
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Banner de Oportunidades */}
                            <div className="space-y-3 bg-[#12121d] p-3 rounded-xl border border-white/5">
                              <span className="text-[10px] font-black text-amber-300 uppercase tracking-wider block">
                                Configuración de Banner de Oportunidades
                              </span>
                              <div className="w-full h-20 sm:h-28 overflow-hidden rounded-lg border border-white/10 bg-black relative">
                                <img
                                  src={settingsDraft.opportunitiesBannerImage || "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=1600&auto=format&fit=crop&q=80"}
                                  alt="Banner Oportunidades"
                                  className="w-full h-full object-cover"
                                />
                              </div>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                <input
                                  type="text"
                                  value={settingsDraft.opportunitiesBannerTitle || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, opportunitiesBannerTitle: e.target.value })}
                                  placeholder="Título: ⏰ OPORTUNIDADES & OFERTAS FLASH"
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-amber-400"
                                />
                                <input
                                  type="text"
                                  value={settingsDraft.opportunitiesBannerSubtitle || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, opportunitiesBannerSubtitle: e.target.value })}
                                  placeholder="Subtítulo: Descuentos por tiempo limitado"
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-amber-400"
                                />
                              </div>
                              <div className="flex gap-2 items-center">
                                <input
                                  type="text"
                                  value={settingsDraft.opportunitiesBannerImage || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, opportunitiesBannerImage: e.target.value })}
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-2 text-white text-xs"
                                  placeholder="URL del banner de Oportunidades"
                                />
                                <label className="bg-amber-500/10 text-amber-300 px-3 py-2 rounded-xl flex items-center cursor-pointer hover:bg-amber-500/20 border border-amber-500/30 text-xs font-bold gap-1">
                                  <Upload size={13} />
                                  <span>Subir</span>
                                  <input
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={(e) =>
                                      handleImageUpload(e, (url) => setSettingsDraft({ ...settingsDraft, opportunitiesBannerImage: url }))
                                    }
                                  />
                                </label>
                              </div>
                            </div>

                            {/* Lista visual de productos seleccionados para Oportunidades */}
                            <div>
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block mb-2">
                                Productos elegidos para Oportunidades ({(settingsDraft.opportunitiesProductIds ?? []).length} de cualquier cantidad):
                              </span>
                              {(settingsDraft.opportunitiesProductIds ?? []).length === 0 ? (
                                <div className="text-center py-4 border border-dashed border-white/10 rounded-xl text-gray-400 text-xs">
                                  Sin selección manual. (Se mostrarán automáticamente productos en oferta).
                                </div>
                              ) : (
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
                                  {(settingsDraft.opportunitiesProductIds ?? []).map((id, idx) => {
                                    const prod = products.find((p) => p.id === id);
                                    if (!prod) return null;
                                    return (
                                      <div
                                        key={id}
                                        className="bg-[#12121d] border border-white/10 rounded-xl p-2 flex items-center justify-between gap-2 shadow"
                                      >
                                        <div className="flex items-center gap-2 min-w-0">
                                          <span className="text-[9px] font-black text-amber-400 px-1 py-0.5 bg-black/40 rounded">
                                            #{idx + 1}
                                          </span>
                                          <p className="text-xs font-bold text-white truncate">{prod.name}</p>
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          <button
                                            type="button"
                                            onClick={() => {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                opportunitiesProductIds: (p.opportunitiesProductIds ?? []).filter((x) => x !== id),
                                              }));
                                            }}
                                            className="p-1 text-red-400 hover:text-red-300"
                                            title="Quitar"
                                          >
                                            <X size={13} />
                                          </button>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>

                            {/* Buscador de catálogo para añadir a Oportunidades */}
                            <div className="pt-2 border-t border-white/5 space-y-2">
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block">
                                Buscar y añadir productos del catálogo a Oportunidades:
                              </span>
                              <input
                                type="text"
                                value={featuredOportunidadesSearch}
                                onChange={(e) => setFeaturedOportunidadesSearch(e.target.value)}
                                placeholder="Buscar por nombre, pasillo..."
                                className="w-full bg-[#12121d] border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:border-amber-400"
                              />
                              <div className="max-h-48 overflow-y-auto space-y-1.5 pr-1">
                                {products
                                  .filter((p) => !p.hidden)
                                  .filter((p) => {
                                    if (!featuredOportunidadesSearch.trim()) return true;
                                    const q = featuredOportunidadesSearch.toLowerCase();
                                    return (
                                      p.name.toLowerCase().includes(q) ||
                                      (p.category && p.category.toLowerCase().includes(q)) ||
                                      (p.aisle && p.aisle.toLowerCase().includes(q))
                                    );
                                  })
                                  .slice(0, 25)
                                  .map((product) => {
                                    const isSelected = (settingsDraft.opportunitiesProductIds ?? []).includes(product.id);
                                    return (
                                      <div
                                        key={product.id}
                                        className={`flex items-center justify-between p-2 rounded-xl border transition-all ${
                                          isSelected
                                            ? "bg-amber-500/10 border-amber-500/30 text-white"
                                            : "bg-[#141420] border-white/5 text-gray-300 hover:border-white/15"
                                        }`}
                                      >
                                        <div className="flex items-center gap-2 min-w-0">
                                          <p className="text-xs font-bold text-white truncate">{product.name}</p>
                                          <span className="text-[10px] text-gray-400 truncate">${Number(product.price || 0).toLocaleString("es-CL")}</span>
                                        </div>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (isSelected) {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                opportunitiesProductIds: (p.opportunitiesProductIds ?? []).filter((x) => x !== product.id),
                                              }));
                                            } else {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                opportunitiesProductIds: [...(p.opportunitiesProductIds ?? []), product.id],
                                              }));
                                              showToast(`Añadido a Oportunidades: "${product.name}"`);
                                            }
                                          }}
                                          className={`px-2.5 py-1 rounded-lg text-[10.5px] font-bold transition-all shrink-0 ${
                                            isSelected
                                              ? "bg-red-500/20 text-red-300 hover:bg-red-500/30"
                                              : "bg-amber-400 text-black hover:bg-amber-300"
                                          }`}
                                        >
                                          {isSelected ? "Quitar" : "+ Añadir"}
                                        </button>
                                      </div>
                                    );
                                  })}
                              </div>
                            </div>
                          </div>

                          {/* BLOQUE D: Pestaña Exclusiva Packs 🎁 (Banner + Selección de Productos) */}
                          <div className="bg-[#181826] p-4 sm:p-5 rounded-2xl border border-white/10 space-y-4">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black text-purple-400 uppercase tracking-wider block">
                                    D. Pestaña Exclusiva "PACKS" 🎁 (Banner + Selección de Productos)
                                  </span>
                                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-purple-500/20 text-purple-300 font-bold border border-purple-500/30">
                                    {(settingsDraft.packsProductIds ?? []).length} seleccionados
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  Configura el banner panorámico y elige la lista de productos (5, 10, 20 o más) que se mostrarán al presionar "PACKS".
                                </p>
                              </div>
                              <div className="flex items-center gap-2 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const packs = products.filter((p) => !p.hidden && (p.name.toLowerCase().includes("pack") || p.category?.toLowerCase().includes("pack") || p.aisle?.toLowerCase().includes("pack"))).map((p) => p.id);
                                    const fallback = products.filter((p) => !p.hidden).map((p) => p.id);
                                    const picked = Array.from(new Set([...packs, ...fallback])).slice(0, 10);
                                    setSettingsDraft((p) => ({ ...p, packsProductIds: picked }));
                                    showToast("Cargados 10 packs/productos");
                                  }}
                                  className="px-3 py-1.5 bg-purple-500/10 hover:bg-purple-500/20 text-purple-300 border border-purple-500/30 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Cargar 10 Packs
                                </button>
                                {(settingsDraft.packsProductIds ?? []).length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setSettingsDraft((p) => ({ ...p, packsProductIds: [] }));
                                      showToast("Lista de Packs limpiada");
                                    }}
                                    className="px-3 py-1.5 bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white rounded-xl text-[10px] font-bold uppercase transition-colors border border-white/10"
                                  >
                                    Limpiar
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Banner de Packs */}
                            <div className="space-y-3 bg-[#12121d] p-3 rounded-xl border border-white/5">
                              <span className="text-[10px] font-black text-purple-300 uppercase tracking-wider block">
                                Configuración de Banner de Packs
                              </span>
                              <div className="w-full h-20 sm:h-28 overflow-hidden rounded-lg border border-white/10 bg-black relative">
                                <img
                                  src={settingsDraft.packsBannerImage || "https://images.unsplash.com/photo-1543007630-9710e4a00a20?w=1600&auto=format&fit=crop&q=80"}
                                  alt="Banner Packs"
                                  className="w-full h-full object-cover"
                                />
                              </div>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                <input
                                  type="text"
                                  value={settingsDraft.packsBannerTitle || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, packsBannerTitle: e.target.value })}
                                  placeholder="Título: 🎁 PACKS & PROMOCIONES"
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-purple-400"
                                />
                                <input
                                  type="text"
                                  value={settingsDraft.packsBannerSubtitle || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, packsBannerSubtitle: e.target.value })}
                                  placeholder="Subtítulo: Arma tu previa con combos"
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-purple-400"
                                />
                              </div>
                              <div className="flex gap-2 items-center">
                                <input
                                  type="text"
                                  value={settingsDraft.packsBannerImage || ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, packsBannerImage: e.target.value })}
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-2 text-white text-xs"
                                  placeholder="URL del banner de Packs"
                                />
                                <label className="bg-purple-500/10 text-purple-300 px-3 py-2 rounded-xl flex items-center cursor-pointer hover:bg-purple-500/20 border border-purple-500/30 text-xs font-bold gap-1">
                                  <Upload size={13} />
                                  <span>Subir</span>
                                  <input
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={(e) =>
                                      handleImageUpload(e, (url) => setSettingsDraft({ ...settingsDraft, packsBannerImage: url }))
                                    }
                                  />
                                </label>
                              </div>
                            </div>

                            {/* Lista visual de productos seleccionados para Packs */}
                            <div>
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block mb-2">
                                Productos elegidos para Packs ({(settingsDraft.packsProductIds ?? []).length} de cualquier cantidad):
                              </span>
                              {(settingsDraft.packsProductIds ?? []).length === 0 ? (
                                <div className="text-center py-4 border border-dashed border-white/10 rounded-xl text-gray-400 text-xs">
                                  Sin selección manual. (Se mostrarán automáticamente productos de la categoría packs).
                                </div>
                              ) : (
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
                                  {(settingsDraft.packsProductIds ?? []).map((id, idx) => {
                                    const prod = products.find((p) => p.id === id);
                                    if (!prod) return null;
                                    return (
                                      <div
                                        key={id}
                                        className="bg-[#12121d] border border-white/10 rounded-xl p-2 flex items-center justify-between gap-2 shadow"
                                      >
                                        <div className="flex items-center gap-2 min-w-0">
                                          <span className="text-[9px] font-black text-purple-400 px-1 py-0.5 bg-black/40 rounded">
                                            #{idx + 1}
                                          </span>
                                          <p className="text-xs font-bold text-white truncate">{prod.name}</p>
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          <button
                                            type="button"
                                            onClick={() => {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                packsProductIds: (p.packsProductIds ?? []).filter((x) => x !== id),
                                              }));
                                            }}
                                            className="p-1 text-red-400 hover:text-red-300"
                                            title="Quitar"
                                          >
                                            <X size={13} />
                                          </button>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>

                            {/* Buscador de catálogo para añadir a Packs */}
                            <div className="pt-2 border-t border-white/5 space-y-2">
                              <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider block">
                                Buscar y añadir productos del catálogo a Packs:
                              </span>
                              <input
                                type="text"
                                value={featuredPacksSearch}
                                onChange={(e) => setFeaturedPacksSearch(e.target.value)}
                                placeholder="Buscar por nombre, pasillo..."
                                className="w-full bg-[#12121d] border border-white/10 rounded-xl px-3 py-2 text-white text-xs focus:border-purple-400"
                              />
                              <div className="max-h-48 overflow-y-auto space-y-1.5 pr-1">
                                {products
                                  .filter((p) => !p.hidden)
                                  .filter((p) => {
                                    if (!featuredPacksSearch.trim()) return true;
                                    const q = featuredPacksSearch.toLowerCase();
                                    return (
                                      p.name.toLowerCase().includes(q) ||
                                      (p.category && p.category.toLowerCase().includes(q)) ||
                                      (p.aisle && p.aisle.toLowerCase().includes(q))
                                    );
                                  })
                                  .slice(0, 25)
                                  .map((product) => {
                                    const isSelected = (settingsDraft.packsProductIds ?? []).includes(product.id);
                                    return (
                                      <div
                                        key={product.id}
                                        className={`flex items-center justify-between p-2 rounded-xl border transition-all ${
                                          isSelected
                                            ? "bg-purple-500/10 border-purple-500/30 text-white"
                                            : "bg-[#141420] border-white/5 text-gray-300 hover:border-white/15"
                                        }`}
                                      >
                                        <div className="flex items-center gap-2 min-w-0">
                                          <p className="text-xs font-bold text-white truncate">{product.name}</p>
                                          <span className="text-[10px] text-gray-400 truncate">${Number(product.price || 0).toLocaleString("es-CL")}</span>
                                        </div>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (isSelected) {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                packsProductIds: (p.packsProductIds ?? []).filter((x) => x !== product.id),
                                              }));
                                            } else {
                                              setSettingsDraft((p) => ({
                                                ...p,
                                                packsProductIds: [...(p.packsProductIds ?? []), product.id],
                                              }));
                                              showToast(`Añadido a Packs: "${product.name}"`);
                                            }
                                          }}
                                          className={`px-2.5 py-1 rounded-lg text-[10.5px] font-bold transition-all shrink-0 ${
                                            isSelected
                                              ? "bg-red-500/20 text-red-300 hover:bg-red-500/30"
                                              : "bg-purple-500 text-white hover:bg-purple-400"
                                          }`}
                                        >
                                          {isSelected ? "Quitar" : "+ Añadir"}
                                        </button>
                                      </div>
                                    );
                                  })}
                              </div>
                            </div>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("delivery");
                                setOpenSettingsSections((p) => ({ ...p, delivery: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Zonas Delivery <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 5: Zonas de Delivery */}
                  {(settingsSubTab === "all" || settingsSubTab === "delivery") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("delivery")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0">
                            <MapPin size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 5: Zonas y Tarifas de Delivery</span>
                              <span className="text-[10px] bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded-full font-bold">
                                Despacho
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Comunas, sectores y valores de costo de envío.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.delivery ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.delivery && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-4 animate-fade-in">
                          <div className="pt-3">
                            <DeliveryLocationsAdminPanel />
                          </div>
                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("hours");
                                setOpenSettingsSections((p) => ({ ...p, hours: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Horarios de Pedido <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 6: Horarios de Atención y Mínimo */}
                  {(settingsSubTab === "all" || settingsSubTab === "hours") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("hours")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center shrink-0">
                            <Clock size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 6: Horario de Recepción y Mínimo</span>
                              <span className="text-[10px] bg-amber-500/15 text-amber-300 border border-amber-500/30 px-2 py-0.5 rounded-full font-bold">
                                {settingsDraft.openTime ?? "11:00"} - {settingsDraft.closeTime ?? "23:00"}
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Control automático de apertura/cierre y monto mínimo para envíos.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.hours ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.hours && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-4 animate-fade-in">
                          <p className="text-xs text-gray-400 pt-3">
                            Fuera de este rango, el sistema desactiva el botón de pedidos y avisa cordialmente a los clientes.
                          </p>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                              <label className="block text-xs font-bold text-amber-400 uppercase mb-2">
                                Hora de Apertura
                              </label>
                              <input
                                type="time"
                                value={(() => {
                                  const mins = parseTimeToMinutes(settingsDraft.openTime || "11:00", 11 * 60);
                                  const h = Math.floor(mins / 60) % 24;
                                  const m = mins % 60;
                                  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
                                })()}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, openTime: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-amber-400 uppercase mb-2">
                                Hora de Cierre
                              </label>
                              <input
                                type="time"
                                value={(() => {
                                  const mins = parseTimeToMinutes(settingsDraft.closeTime || "23:45", 23 * 60 + 45);
                                  const h = Math.floor(mins / 60) % 24;
                                  const m = mins % 60;
                                  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
                                })()}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, closeTime: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                              />
                            </div>
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                              Monto mínimo de compra para Delivery ($)
                            </label>
                            <input
                              type="number"
                              min={0}
                              step={1000}
                              value={settingsDraft.deliveryMinimum ?? 10000}
                              onChange={(e) =>
                                setSettingsDraft({ ...settingsDraft, deliveryMinimum: Number(e.target.value) })
                              }
                              className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                              placeholder="10000"
                            />
                            <p className="text-[11px] text-gray-500 mt-1">Los pedidos con Retiro en Tienda no exigen monto mínimo.</p>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("contingency");
                                setOpenSettingsSections((p) => ({ ...p, contingency: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Modo Contingencia <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 7: Modo Contingencia */}
                  {(settingsSubTab === "all" || settingsSubTab === "contingency") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("contingency")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-red-500/10 border border-red-500/20 text-red-400 flex items-center justify-center shrink-0">
                            <Ban size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 7: Modo Contingencia (Catálogo Reducido)</span>
                              <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold border ${
                                settingsDraft.contingencyMode
                                  ? "bg-red-500/20 text-red-300 border-red-500/40"
                                  : "bg-emerald-500/20 text-emerald-300 border-emerald-500/40"
                              }`}>
                                {settingsDraft.contingencyMode ? "🔴 ACTIVADO" : "🟢 Inactivo"}
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Filtra la tienda para mostrar solo los 30 productos esenciales de entrega inmediata.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.contingency ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.contingency && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-4 animate-fade-in">
                          <div className="pt-3">
                            <label className="flex items-center gap-3 p-4 bg-[#181826] border border-red-500/30 rounded-2xl cursor-pointer hover:border-red-500/60 transition-colors">
                              <input
                                type="checkbox"
                                checked={Boolean(settingsDraft.contingencyMode)}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, contingencyMode: e.target.checked })
                                }
                                className="w-5 h-5 accent-red-500 rounded cursor-pointer"
                              />
                              <div>
                                <span className="text-xs font-black text-white uppercase tracking-wider block">
                                  Habilitar Modo Contingencia
                                </span>
                                <span className="text-[11px] text-gray-400">
                                  {settingsDraft.contingencyMode
                                    ? "🔴 MODO CONTINGENCIA ACTIVO (La tienda recibe pedidos solo del catálogo seleccionado)"
                                    : "🟢 Estado normal (Todos los productos del catálogo visibles)"}
                                </span>
                              </div>
                            </label>
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-red-400 uppercase mb-2">
                              Mensaje del Banner de Emergencia
                            </label>
                            <textarea
                              rows={2}
                              value={settingsDraft.contingencyMessage ?? ""}
                              onChange={(e) =>
                                setSettingsDraft({ ...settingsDraft, contingencyMessage: e.target.value })
                              }
                              className="w-full bg-[#181826] border border-red-500/30 rounded-xl p-3 text-white text-xs focus:border-red-400 focus:outline-none"
                              placeholder="🚨 MODO CONTINGENCIA: Mostrando catálogo reducido de disponibilidad inmediata."
                            />
                          </div>

                          {/* Selector de 30 productos */}
                          <div className="bg-[#181826] p-4 rounded-2xl border border-white/10 space-y-3">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div>
                                <span className="text-xs font-black text-white uppercase tracking-wider block">
                                  Catálogo de Contingencia Seleccionado
                                </span>
                                <span className="text-[11px] font-mono font-bold text-amber-400">
                                  {products.filter((p) => p.contingencyEnabled).length} / 30 Productos seleccionados
                                </span>
                              </div>
                              <div className="flex gap-2">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const first30Ids = products.slice(0, 30).map((p) => p.id);
                                    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                      old
                                        ? {
                                            ...old,
                                            products: old.products.map((p: any, i: number) => ({
                                              ...p,
                                              contingencyEnabled: i < 30,
                                            })),
                                          }
                                        : old,
                                    );
                                    showToast("Seleccionados primeros 30 productos para Contingencia");
                                    fetch(`${import.meta.env.BASE_URL}api/admin/products/contingency-batch`, {
                                      method: "POST",
                                      headers: { "Content-Type": "application/json" },
                                      body: JSON.stringify({ ids: first30Ids, enable: true }),
                                    }).catch(() => refreshMenu());
                                  }}
                                  className="px-3 py-1.5 bg-red-500/10 border border-red-500/30 text-red-300 hover:bg-red-500/20 rounded-xl text-[10px] font-bold uppercase transition-colors"
                                >
                                  ⚡ Seleccionar Primeros 30
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                      old
                                        ? {
                                            ...old,
                                            products: old.products.map((p: any) => ({
                                              ...p,
                                              contingencyEnabled: false,
                                            })),
                                          }
                                        : old,
                                    );
                                    showToast("Catálogo de contingencia limpiado");
                                    fetch(`${import.meta.env.BASE_URL}api/admin/products/contingency-batch`, {
                                      method: "POST",
                                      headers: { "Content-Type": "application/json" },
                                      body: JSON.stringify({ ids: products.map((p) => p.id), enable: false }),
                                    }).catch(() => refreshMenu());
                                  }}
                                  className="px-3 py-1.5 bg-white/5 text-gray-400 hover:text-white rounded-xl text-[10px] font-bold uppercase transition-colors border border-white/5"
                                >
                                  Deseleccionar Todos
                                </button>
                              </div>
                            </div>

                            <input
                              type="text"
                              value={adminProductSearch}
                              onChange={(e) => setAdminProductSearch(e.target.value)}
                              placeholder="Buscar producto para activar en contingencia..."
                              className="w-full bg-[#12121d] border border-white/10 rounded-xl px-3.5 py-2.5 text-white text-xs focus:border-red-400 focus:outline-none"
                            />

                            <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
                              {products
                                .filter((p) => !adminProductSearch || p.name.toLowerCase().includes(adminProductSearch.toLowerCase()))
                                .map((product) => (
                                  <label
                                    key={product.id}
                                    className={`flex items-center justify-between p-2 rounded-xl border transition-all cursor-pointer ${
                                      product.contingencyEnabled
                                        ? "bg-red-950/20 border-red-500/40 text-white"
                                        : "bg-[#12121d] border-white/5 text-gray-400 hover:border-white/10"
                                    }`}
                                  >
                                    <div className="flex items-center gap-2.5 min-w-0">
                                      <input
                                        type="checkbox"
                                        checked={Boolean(product.contingencyEnabled)}
                                        onChange={() => {
                                          const nextVal = !product.contingencyEnabled;
                                          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                            old
                                              ? {
                                                  ...old,
                                                  products: old.products.map((p: any) =>
                                                    p.id === product.id ? { ...p, contingencyEnabled: nextVal } : p,
                                                  ),
                                                }
                                              : old,
                                          );
                                          fetch(`${import.meta.env.BASE_URL}api/products/${product.id}`, {
                                            method: "PATCH",
                                            headers: { "Content-Type": "application/json" },
                                            body: JSON.stringify({ contingencyEnabled: nextVal }),
                                          }).catch(() => refreshMenu());
                                        }}
                                        className="w-4 h-4 accent-red-500 rounded cursor-pointer shrink-0"
                                      />
                                      {product.image && (
                                        <img src={product.image} alt="" className="w-7 h-7 object-cover rounded shrink-0 bg-black" />
                                      )}
                                      <div className="truncate text-xs">
                                        <span className="font-bold block truncate">{product.name}</span>
                                        <span className="text-[10px] text-gray-500">${product.price.toLocaleString("es-CL")} · {product.category}</span>
                                      </div>
                                    </div>
                                    {product.contingencyEnabled && (
                                      <span className="text-[9px] font-black uppercase tracking-wider px-2 py-0.5 bg-red-500/20 text-red-300 rounded border border-red-500/30 shrink-0">
                                        🚨 Seleccionado
                                      </span>
                                    )}
                                  </label>
                                ))}
                            </div>
                          </div>

                          <div className="flex justify-end pt-2">
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsSubTab("contact");
                                setOpenSettingsSections((p) => ({ ...p, contact: true }));
                              }}
                              className="px-4 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                            >
                              Siguiente: Contacto & Redes <ChevronRight size={14} />
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 8: Contacto, Redes & Créditos */}
                  {(settingsSubTab === "all" || settingsSubTab === "contact") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("contact")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-blue-500/10 border border-blue-500/20 text-blue-400 flex items-center justify-center shrink-0">
                            <Phone size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 8: Contacto, WhatsApp & Créditos</span>
                              <span className="text-[10px] bg-blue-500/15 text-blue-300 border border-blue-500/30 px-2 py-0.5 rounded-full font-bold">
                                {settingsDraft.contactPhone ? "Listo" : "Incompleto"}
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Número de pedidos WhatsApp, dirección física, horarios y pie de página.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          {openSettingsSections.contact ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.contact && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-4 animate-fade-in">
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4">
                            <div>
                              <label className="block text-xs font-bold text-[#ffd025] uppercase mb-2">
                                Número WhatsApp para Pedidos
                              </label>
                              <input
                                type="text"
                                value={settingsDraft.whatsapp || settingsDraft.contactPhone || ""}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, whatsapp: e.target.value, contactPhone: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] font-mono text-sm placeholder-gray-500"
                                placeholder="Ej: 56912345678"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                                Dirección / Ubicación
                              </label>
                              <input
                                type="text"
                                value={settingsDraft.contactAddress ?? ""}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, contactAddress: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm placeholder-gray-500"
                                placeholder="Ej: Av. Principal 123, Santiago"
                              />
                            </div>
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-gray-300 uppercase mb-2">
                              Texto de Horarios para el Pie de Página
                            </label>
                            <textarea
                              rows={2}
                              value={settingsDraft.contactHours ?? ""}
                              onChange={(e) =>
                                setSettingsDraft({ ...settingsDraft, contactHours: e.target.value })
                              }
                              className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white text-xs focus:border-[#ffd025]"
                              placeholder={"Lun-Vie: 10:00 - 22:00\nSáb-Dom: 11:00 - 23:00"}
                            />
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-4 border-t border-white/5">
                            <div>
                              <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                                Logo del Pie de Página (URL)
                              </label>
                              <div className="flex gap-2">
                                <input
                                  type="text"
                                  value={settingsDraft.footerLogo ?? ""}
                                  onChange={(e) =>
                                    setSettingsDraft({ ...settingsDraft, footerLogo: e.target.value })
                                  }
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                                  placeholder="URL del logo"
                                />
                                <label className="bg-[#ffd025]/10 text-[#ffd025] px-3.5 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                                  <Upload size={18} />
                                  <input
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={(e) =>
                                      handleImageUpload(e, (url) =>
                                        setSettingsDraft({ ...settingsDraft, footerLogo: url }),
                                      )
                                    }
                                  />
                                </label>
                              </div>
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                                Pequeña Descripción para el Pie de Página
                              </label>
                              <input
                                type="text"
                                value={settingsDraft.footerDescription ?? ""}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, footerDescription: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                                placeholder="Tu botillería de confianza..."
                              />
                            </div>
                          </div>

                          <div className="pt-3 border-t border-white/5 space-y-3">
                            <h4 className="text-xs font-black uppercase text-[#ffd025] tracking-wider">Redes Sociales (URLs)</h4>
                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                              <div>
                                <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Instagram URL</label>
                                <input
                                  type="text"
                                  value={settingsDraft.socialInstagram ?? ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, socialInstagram: e.target.value })}
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                  placeholder="https://instagram.com/..."
                                />
                              </div>
                              <div>
                                <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Facebook URL</label>
                                <input
                                  type="text"
                                  value={settingsDraft.socialFacebook ?? ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, socialFacebook: e.target.value })}
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                  placeholder="https://facebook.com/..."
                                />
                              </div>
                              <div>
                                <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">TikTok URL</label>
                                <input
                                  type="text"
                                  value={settingsDraft.socialTiktok ?? ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, socialTiktok: e.target.value })}
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                  placeholder="https://tiktok.com/@..."
                                />
                              </div>
                              <div>
                                <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">WhatsApp URL / Número</label>
                                <input
                                  type="text"
                                  value={settingsDraft.socialWhatsapp ?? ""}
                                  onChange={(e) => setSettingsDraft({ ...settingsDraft, socialWhatsapp: e.target.value })}
                                  className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025]"
                                  placeholder="56912345678"
                                />
                              </div>
                            </div>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2 border-t border-white/5">
                            <div>
                              <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                                Nombre de Agencia / Desarrollador
                              </label>
                              <input
                                type="text"
                                value={settingsDraft.agencyName}
                                onChange={(e) =>
                                  setSettingsDraft({ ...settingsDraft, agencyName: e.target.value })
                                }
                                className="w-full bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                                placeholder="Ej: Web Studio"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                                Logo de Agencia (Opcional)
                              </label>
                              <div className="flex gap-2">
                                <input
                                  type="text"
                                  value={settingsDraft.agencyLogo}
                                  onChange={(e) =>
                                    setSettingsDraft({ ...settingsDraft, agencyLogo: e.target.value })
                                  }
                                  onBlur={(e) =>
                                    resolveImageUrl(e.target.value, (r) =>
                                      setSettingsDraft((p) => ({ ...p, agencyLogo: r })),
                                    )
                                  }
                                  className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                                  placeholder="URL del logo"
                                />
                                <label className="bg-[#ffd025]/10 text-[#ffd025] px-3.5 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                                  <Upload size={18} />
                                  <input
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={(e) =>
                                      handleImageUpload(e, (url) =>
                                        setSettingsDraft({ ...settingsDraft, agencyLogo: url }),
                                      )
                                    }
                                  />
                                </label>
                              </div>
                            </div>
                          </div>

                          {/* Sistema Anti-Inactividad Keep-Alive Activo */}
                          <div className="pt-4 border-t border-white/5">
                            <div className="bg-[#12121d] border border-emerald-500/30 rounded-2xl p-4 flex items-center gap-3.5">
                              <div className="w-9 h-9 rounded-xl bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shrink-0">
                                <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 animate-ping" />
                              </div>
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black uppercase text-emerald-300">
                                    Sistema Keep-Alive 24/7 Activo
                                  </span>
                                  <span className="text-[9.5px] bg-emerald-500/20 text-emerald-300 px-2 py-0.5 rounded-full font-bold">
                                    Anti-Hibernación Render
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5 leading-snug">
                                  El servidor ejecuta un pulso continuo cada 4 minutos y auto-sincroniza la base de datos en disco para evitar que la página se suspenda o pierda información en planes gratuitos.
                                </p>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 9: Galería de 6 Imágenes (Aspecto 4:5) */}
                  {(settingsSubTab === "all" || settingsSubTab === "gallery") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("gallery")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center shrink-0">
                            <Camera size={18} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 9: Galería de Fotos (Aspecto 4:5)</span>
                              <span className="text-[10px] bg-amber-500/15 text-amber-300 border border-amber-500/30 px-2 py-0.5 rounded-full font-bold">
                                6 Fotos en PC / 2 en Celular
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">
                              Personaliza las 6 fotos con formato vertical 4:5 que se muestran antes del pie de página.
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          <span className="text-xs font-mono font-bold hidden sm:inline">
                            {openSettingsSections.gallery ? "Plegar" : "Editar"}
                          </span>
                          {openSettingsSections.gallery ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                        </div>
                      </button>

                      {openSettingsSections.gallery && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-6 animate-fade-in">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-3">
                            <div>
                              <h5 className="text-xs font-black uppercase text-[#ffd025] tracking-wider flex items-center gap-1.5">
                                <span>📸</span> 6 Fotos en Formato 4:5 (Sin Recuadro ni Puntas Redondeadas)
                              </h5>
                              <p className="text-[11px] text-gray-400 mt-0.5">
                                En celular se muestran las fotos 1 y 2. En PC se muestran las 6 fotos completas sueltas en la grilla.
                              </p>
                            </div>
                            <button
                              type="button"
                              onClick={() => {
                                setSettingsDraft({
                                  ...settingsDraft,
                                  galleryImages: [...DEFAULT_GALLERY_IMAGES]
                                });
                                showToast("Fotos por defecto restauradas");
                              }}
                              className="px-3 py-1.5 bg-white/5 hover:bg-white/10 border border-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 self-start sm:self-auto shrink-0"
                            >
                              ⚡ Restaurar Fotos por Defecto
                            </button>
                          </div>

                          {/* 6 Ranuras de Imagen */}
                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                            {Array.from({ length: 6 }).map((_, slotIdx) => {
                              const currentImages = settingsDraft.galleryImages || DEFAULT_GALLERY_IMAGES;
                              const currentItem: GalleryImageItem = currentImages[slotIdx] || {
                                id: `gal-${slotIdx + 1}`,
                                url: "",
                                title: "",
                                aisle: "",
                                buttonText: "Ver más"
                              };

                              const updateSlotField = (field: keyof GalleryImageItem, val: string) => {
                                const copy = [...(settingsDraft.galleryImages || DEFAULT_GALLERY_IMAGES)];
                                copy[slotIdx] = {
                                  ...(copy[slotIdx] || { id: `gal-${slotIdx + 1}`, url: "" }),
                                  [field]: val
                                };
                                setSettingsDraft({ ...settingsDraft, galleryImages: copy });
                              };

                              const isMobileSlot = slotIdx < 2;

                              return (
                                <div
                                  key={slotIdx}
                                  className="bg-[#181826] border border-white/10 rounded-2xl p-4 flex flex-col justify-between space-y-3 relative group"
                                >
                                  {/* Badge de Ranura */}
                                  <div className="flex items-center justify-between">
                                    <div className="flex items-center gap-2">
                                      <span className="text-xs font-black text-[#ffd025] bg-[#ffd025]/10 border border-[#ffd025]/20 px-2.5 py-0.5 rounded-lg">
                                        Foto #{slotIdx + 1}
                                      </span>
                                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${
                                        isMobileSlot
                                          ? "bg-emerald-500/15 text-emerald-300 border border-emerald-500/30"
                                          : "bg-blue-500/15 text-blue-300 border border-blue-500/30"
                                      }`}>
                                        {isMobileSlot ? "📱 Celular & PC" : "💻 Solo en PC"}
                                      </span>
                                    </div>
                                    <span className="text-[10px] font-mono text-gray-500">
                                      Aspecto 4:5
                                    </span>
                                  </div>

                                  {/* Previsualización en vivo Aspecto 4:5 sin textos ni botones */}
                                  <div className="relative aspect-[4/5] w-full overflow-hidden bg-black/60 border border-white/10 group-hover:border-[#ffd025]/30 transition-all">
                                    {currentItem.url ? (
                                      <img
                                        src={currentItem.url}
                                        alt={`Foto ${slotIdx + 1}`}
                                        className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-105"
                                      />
                                    ) : (
                                      <div className="w-full h-full flex flex-col items-center justify-center text-gray-500 gap-2 p-4 text-center">
                                        <Camera size={28} className="opacity-40" />
                                        <span className="text-xs">Sin imagen (Formato 4:5)</span>
                                      </div>
                                    )}
                                    <div className="absolute top-2 left-2 px-1.5 py-0.5 bg-black/80 text-[9px] font-mono text-[#ffd025] border border-white/15">
                                      4:5
                                    </div>
                                  </div>

                                  {/* Inputs de URL & Subir archivo */}
                                  <div>
                                    <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                                      URL de Imagen o Subir Archivo
                                    </label>
                                    <div className="flex gap-2">
                                      <input
                                        type="text"
                                        value={currentItem.url || ""}
                                        onChange={(e) => updateSlotField("url", e.target.value)}
                                        className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs focus:border-[#ffd025]"
                                        placeholder="https://images.unsplash.com/..."
                                      />
                                      <label
                                        className="bg-[#ffd025]/10 text-[#ffd025] px-3 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30 shrink-0"
                                        title="Subir foto desde dispositivo"
                                      >
                                        <Upload size={14} />
                                        <input
                                          type="file"
                                          accept="image/*"
                                          className="hidden"
                                          onChange={(e) =>
                                            handleImageUpload(e, (url) => updateSlotField("url", url))
                                          }
                                        />
                                      </label>
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* SECCIÓN 10: Persistencia, Copias de Seguridad & Render */}
                  {(settingsSubTab === "all" || settingsSubTab === "backup") && (
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 overflow-hidden shadow-xl shadow-black/50 transition-all">
                      <button
                        type="button"
                        onClick={() => toggleSettingsSection("backup")}
                        className="w-full p-5 sm:p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0">
                            <ShieldCheck size={20} />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-black text-white text-base uppercase">Paso 10: Copias de Seguridad & Persistencia (Render / GitHub)</span>
                              <span className="text-[10px] bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded-full font-bold">
                                Anti-Pérdida
                              </span>
                            </div>
                            <p className="text-xs text-gray-400 mt-0.5">Protege tus productos, fotos y configuraciones ante cualquier actualización o redespliegue.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 text-gray-400">
                          <span className="text-xs font-mono font-bold hidden sm:inline">
                            {openSettingsSections.backup ? "Plegar" : "Abrir"}
                          </span>
                          {openSettingsSections.backup ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                        </div>
                      </button>

                      {openSettingsSections.backup && (
                        <div className="p-6 pt-0 border-t border-white/5 space-y-6 animate-fade-in">
                          {/* Banner explicativo paso a paso */}
                          <div className="bg-gradient-to-br from-emerald-950/40 via-[#181826] to-[#12121d] border border-emerald-500/30 rounded-2xl p-5 mt-4">
                            <div className="flex items-start gap-3.5">
                              <div className="w-9 h-9 rounded-xl bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shrink-0 mt-0.5">
                                <Archive size={18} />
                              </div>
                              <div className="space-y-2 text-xs">
                                <h4 className="font-black text-sm text-white uppercase tracking-wider flex items-center gap-2">
                                  ¿Cómo mantener todas tus fotos y productos al actualizar en Render o GitHub?
                                </h4>
                                <p className="text-gray-300 leading-relaxed">
                                  En servicios como Render (en planes gratuitos sin disco persistente), cada despliegue desde GitHub crea un contenedor nuevo. Con las siguientes 2 opciones nunca más perderás tus imágenes ni productos:
                                </p>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-2">
                                  <div className="bg-black/30 p-3 rounded-xl border border-white/10 space-y-1">
                                    <span className="font-bold text-[#ffd025] flex items-center gap-1.5 text-[11px] uppercase">
                                      <DownloadCloud size={14} /> Opción 1: Respaldo ZIP (1 Clic)
                                    </span>
                                    <p className="text-gray-400 text-[11px]">
                                      Haz clic en <strong>Descargar Respaldo ZIP</strong> y guárdalo en tu computador. Contiene todos tus productos, precios, fotos y ajustes. Al actualizar la página, haces clic en <strong>Restaurar Copia</strong> y en 2 segundos todo vuelve a estar como antes.
                                    </p>
                                  </div>
                                  <div className="bg-black/30 p-3 rounded-xl border border-white/10 space-y-1">
                                    <span className="font-bold text-emerald-400 flex items-center gap-1.5 text-[11px] uppercase">
                                      <Database size={14} /> Opción 2: Auto-Sincronización en Git
                                    </span>
                                    <p className="text-gray-400 text-[11px]">
                                      Cada imagen que subes se auto-guarda en <code className="text-[#ffd025]">src/server/seed_uploads/</code> y tus productos en <code className="text-[#ffd025]">src/server/seed_db.json</code>. Si haces commit de esos archivos a tu repositorio Git, Render los compilará automáticamente en cada despliegue.
                                    </p>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* Botones de Acción Inmediata */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                            {/* 1. Descargar Respaldo ZIP */}
                            <button
                              type="button"
                              onClick={handleExportBackupZip}
                              className="p-4 bg-gradient-to-br from-purple-900/30 to-purple-950/50 hover:from-purple-900/50 hover:to-purple-950/70 border border-purple-500/40 rounded-2xl flex flex-col justify-between text-left transition-all hover:scale-[1.01] active:scale-[0.99] group shadow-lg shadow-purple-950/50 cursor-pointer"
                            >
                              <div className="flex items-center justify-between mb-3">
                                <div className="w-10 h-10 rounded-xl bg-purple-500/20 text-purple-300 flex items-center justify-center border border-purple-500/30 group-hover:scale-110 transition-transform">
                                  <Archive size={20} />
                                </div>
                                <span className="text-[10px] uppercase font-mono font-bold px-2 py-0.5 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/30">
                                  Recomendado
                                </span>
                              </div>
                              <div>
                                <h5 className="font-black text-white text-sm uppercase">Descargar Respaldo ZIP</h5>
                                <p className="text-[11px] text-gray-400 mt-1">Descarga un archivo .ZIP con la base de datos completa y TODAS las fotos subidas.</p>
                              </div>
                            </button>

                            {/* 2. Restaurar Respaldo ZIP o JSON */}
                            <label className="p-4 bg-gradient-to-br from-emerald-900/30 to-emerald-950/50 hover:from-emerald-900/50 hover:to-emerald-950/70 border border-emerald-500/40 rounded-2xl flex flex-col justify-between text-left transition-all hover:scale-[1.01] active:scale-[0.99] group shadow-lg shadow-emerald-950/50 cursor-pointer">
                              <input
                                type="file"
                                accept=".zip, .json"
                                onChange={handleRestoreBackup}
                                className="hidden"
                              />
                              <div className="flex items-center justify-between mb-3">
                                <div className="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-300 flex items-center justify-center border border-emerald-500/30 group-hover:scale-110 transition-transform">
                                  <UploadCloud size={20} />
                                </div>
                                <span className="text-[10px] uppercase font-mono font-bold px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                  1-Clic
                                </span>
                              </div>
                              <div>
                                <h5 className="font-black text-white text-sm uppercase">Restaurar Copia (.ZIP / .JSON)</h5>
                                <p className="text-[11px] text-gray-400 mt-1">Sube tu archivo .ZIP o .JSON para restaurar todos tus productos y fotos inmediatamente.</p>
                              </div>
                            </label>

                            {/* 3. Exportar Catálogo JSON */}
                            <button
                              type="button"
                              onClick={handleExportBackupJson}
                              className="p-4 bg-gradient-to-br from-blue-900/30 to-blue-950/50 hover:from-blue-900/50 hover:to-blue-950/70 border border-blue-500/40 rounded-2xl flex flex-col justify-between text-left transition-all hover:scale-[1.01] active:scale-[0.99] group shadow-lg shadow-blue-950/50 cursor-pointer"
                            >
                              <div className="flex items-center justify-between mb-3">
                                <div className="w-10 h-10 rounded-xl bg-blue-500/20 text-blue-300 flex items-center justify-center border border-blue-500/30 group-hover:scale-110 transition-transform">
                                  <Download size={20} />
                                </div>
                                <span className="text-[10px] uppercase font-mono font-bold px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-300 border border-blue-500/30">
                                  Liviano
                                </span>
                              </div>
                              <div>
                                <h5 className="font-black text-white text-sm uppercase">Descargar Catálogo JSON</h5>
                                <p className="text-[11px] text-gray-400 mt-1">Copia de texto con todos los productos, categorías, pedidos y configuración.</p>
                              </div>
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {/* Final floating save action */}
                <div className="pt-2">
                  <button
                    type="button"
                    onClick={saveSettings}
                    className="w-full py-4 bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] rounded-2xl font-black uppercase text-sm tracking-wider hover:scale-[1.005] active:scale-[0.99] transition-all shadow-xl shadow-[#ffd025]/20 flex items-center justify-center gap-2"
                  >
                    <CheckCircle size={18} /> Guardar Todos los Ajustes a la Nube
                  </button>
                </div>
              </div>
            )}

            {adminTab === "classifications" && (
              <div className="space-y-6 max-w-5xl mx-auto animate-fade-in">
                {/* Header with step overview and action */}
                <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-7 rounded-3xl border border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-xl shadow-black/60">
                  <div>
                    <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
                      <Tag size={15} /> Arquitectura del Catálogo
                    </div>
                    <h2 className="text-xl sm:text-2xl font-black text-white uppercase">Clasificaciones & Pasillos</h2>
                    <p className="text-xs text-gray-400 mt-1">Configura paso a paso o usa los menús plegables para ordenar categorías, pasillos y etiquetas.</p>
                  </div>
                  <div className="flex items-center gap-2.5">
                    <button
                      type="button"
                      onClick={() => {
                        const allOpen = Object.values(openClassificationsSections).every(Boolean);
                        const nextVal = !allOpen;
                        setOpenClassificationsSections({
                          categories: nextVal,
                          aisles: nextVal,
                          subcategories: nextVal,
                        });
                      }}
                      className="px-3.5 py-2.5 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold transition-all border border-white/10 flex items-center gap-2"
                    >
                      <Layers size={14} />
                      {Object.values(openClassificationsSections).every(Boolean) ? "Plegar Todos" : "Desplegar Todos"}
                    </button>
                  </div>
                </div>

                {/* Step Sub-Tabs Navigation */}
                <div className="flex items-center gap-2 overflow-x-auto pb-2 border-b border-white/10 scrollbar-hide">
                  {[
                    { id: "categories", label: "Paso 1: Categorías", icon: Layers, count: categoriesData.length, hint: "Pestañas superiores" },
                    { id: "aisles", label: "Paso 2: Pasillos & IA", icon: Tag, count: aislesData.length, hint: "Separadores & Banners" },
                    { id: "subcategories", label: "Paso 3: Subcategorías", icon: CheckCircle, count: subcategoriesData.length, hint: "Etiquetas y tags" },
                    { id: "all", label: "Ver Todas en Plegables", icon: Layers, count: categoriesData.length + aislesData.length + subcategoriesData.length, hint: "Vista global" },
                  ].map((tab) => {
                    const Icon = tab.icon;
                    const isActive = classificationsSubTab === tab.id;
                    return (
                      <button
                        key={tab.id}
                        type="button"
                        onClick={() => {
                          setClassificationsSubTab(tab.id as any);
                          if (tab.id !== "all") {
                            setOpenClassificationsSections((prev) => ({ ...prev, [tab.id]: true }));
                          }
                        }}
                        className={`flex items-center gap-2.5 px-4 py-2.5 rounded-2xl text-xs font-bold transition-all whitespace-nowrap shrink-0 border ${
                          isActive
                            ? "bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#0a0a0f] border-[#ffd025] shadow-lg shadow-[#ffd025]/20 font-black"
                            : "bg-[#141420]/80 text-gray-400 hover:text-white border-white/5 hover:border-white/20"
                        }`}
                      >
                        <Icon size={14} className={isActive ? "text-[#0a0a0f]" : "text-[#ffd025]"} />
                        <span>{tab.label}</span>
                        <span
                          className={`text-[10px] px-2 py-0.5 rounded-full font-black ${
                            isActive ? "bg-black/20 text-[#0a0a0f]" : "bg-white/10 text-gray-300"
                          }`}
                        >
                          {tab.count}
                        </span>
                      </button>
                    );
                  })}
                </div>

                {/* Section 1: Categorías */}
                {(classificationsSubTab === "categories" || classificationsSubTab === "all") && (
                  <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 shadow-2xl shadow-black/80 overflow-hidden">
                    <button
                      type="button"
                      onClick={() =>
                        setOpenClassificationsSections((p) => ({ ...p, categories: !p.categories }))
                      }
                      className="w-full p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors border-b border-white/5"
                    >
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-2xl bg-[#ffd025]/10 border border-[#ffd025]/20 text-[#ffd025] flex items-center justify-center font-black">
                          1
                        </div>
                        <div>
                          <h3 className="text-lg font-black uppercase text-white flex items-center gap-2">
                            <Layers size={18} className="text-[#ffd025]" /> Categorías Principales
                            <span className="text-xs font-bold text-[#ffd025] bg-[#ffd025]/10 px-2 py-0.5 rounded-full border border-[#ffd025]/20">
                              {categoriesData.length} configuradas
                            </span>
                          </h3>
                          <p className="text-xs text-gray-400">Pestañas superiores de la tienda. Arrastra para reordenar.</p>
                        </div>
                      </div>
                      <div className="p-2 rounded-xl bg-white/5 text-gray-400">
                        {openClassificationsSections.categories ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                      </div>
                    </button>

                    {openClassificationsSections.categories && (
                      <div className="p-6 space-y-5 animate-fade-in">
                        <div className="flex gap-2.5">
                          <input
                            type="text"
                            value={newCategoryName}
                            onChange={(e) => setNewCategoryName(e.target.value)}
                            placeholder="Nueva categoría (ej: Cervezas, Vinos, Destilados)..."
                            onKeyDown={(e) => e.key === "Enter" && handleAddCategory()}
                            className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3.5 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                          />
                          <button
                            onClick={handleAddCategory}
                            className="px-5 py-3.5 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] rounded-xl font-black text-sm uppercase transition-all shadow-md hover:scale-[1.02] flex items-center gap-1.5"
                          >
                            <Plus size={18} /> Añadir
                          </button>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5">
                          {categoriesData.map((cat, index) => (
                            <div
                              key={cat.id}
                              draggable
                              onDragStart={(e) => handleDragStart(e, index)}
                              onDragOver={(e) => e.preventDefault()}
                              onDrop={(e) => handleDrop(e, index)}
                              className="flex justify-between items-center bg-[#181826] p-3.5 rounded-xl border border-white/5 cursor-move hover:border-[#ffd025]/50 transition-colors group"
                            >
                              <div className="flex items-center gap-3">
                                <GripVertical size={16} className="text-gray-500 group-hover:text-[#ffd025]" />
                                <span className="font-bold text-white text-sm">{cat.name}</span>
                              </div>
                              <button
                                onClick={() => deleteCategoryHandler(cat.id)}
                                className="text-red-400/60 hover:text-red-400 p-1.5 rounded-lg hover:bg-red-500/10 transition-colors"
                                title="Eliminar categoría"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          ))}
                        </div>

                        <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-3 border-t border-white/5">
                          <button
                            onClick={() =>
                              reorderCategoriesMut.mutate(
                                { data: { ids: categoriesData.map((c) => c.id) } },
                                {
                                  onSuccess: () => {
                                    setSavedCats(true);
                                    setTimeout(() => setSavedCats(false), 2500);
                                  },
                                },
                              )
                            }
                            disabled={reorderCategoriesMut.isPending}
                            className="w-full sm:w-auto px-6 py-3 rounded-xl font-black uppercase text-xs transition-all flex items-center justify-center gap-2 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] hover:scale-[1.01] disabled:opacity-60 shadow-lg shadow-[#ffd025]/20"
                          >
                            {savedCats ? (
                              <><CheckCircle size={15} /> Categorías Guardadas</>
                            ) : reorderCategoriesMut.isPending ? (
                              "Guardando..."
                            ) : (
                              "Guardar Orden de Categorías"
                            )}
                          </button>

                          {classificationsSubTab === "categories" && (
                            <button
                              type="button"
                              onClick={() => {
                                setClassificationsSubTab("aisles");
                                setOpenClassificationsSections((p) => ({ ...p, aisles: true }));
                              }}
                              className="w-full sm:w-auto px-5 py-3 rounded-xl font-black uppercase text-xs bg-white/5 hover:bg-white/10 text-white border border-white/10 transition-all flex items-center justify-center gap-2"
                            >
                              Siguiente: Configurar Pasillos ➔
                            </button>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* Section 2: Pasillos & Banners con IA */}
                {(classificationsSubTab === "aisles" || classificationsSubTab === "all") && (
                  <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 shadow-2xl shadow-black/80 overflow-hidden">
                    <button
                      type="button"
                      onClick={() =>
                        setOpenClassificationsSections((p) => ({ ...p, aisles: !p.aisles }))
                      }
                      className="w-full p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors border-b border-white/5"
                    >
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-2xl bg-[#ffd025]/10 border border-[#ffd025]/20 text-[#ffd025] flex items-center justify-center font-black">
                          2
                        </div>
                        <div>
                          <h3 className="text-lg font-black uppercase text-white flex items-center gap-2">
                            <Tag size={18} className="text-[#ffd025]" /> Pasillos & Banners Inteligentes
                            <span className="text-xs font-bold text-[#ffd025] bg-[#ffd025]/10 px-2 py-0.5 rounded-full border border-[#ffd025]/20">
                              {aislesData.length} pasillos
                            </span>
                          </h3>
                          <p className="text-xs text-gray-400">Separadores visuales con soporte para banners promocionales generados con IA.</p>
                        </div>
                      </div>
                      <div className="p-2 rounded-xl bg-white/5 text-gray-400">
                        {openClassificationsSections.aisles ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                      </div>
                    </button>

                    {openClassificationsSections.aisles && (
                      <div className="p-6 space-y-5 animate-fade-in">
                        <div className="flex gap-2.5">
                          <input
                            type="text"
                            value={newAisleName}
                            onChange={(e) => setNewAisleName(e.target.value)}
                            placeholder="Nuevo pasillo (ej: Cervezas Artesanales, Piscos Especiales)..."
                            onKeyDown={(e) => e.key === "Enter" && handleAddAisle()}
                            className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3.5 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                          />
                          <button
                            onClick={handleAddAisle}
                            className="px-5 py-3.5 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] rounded-xl font-black text-sm uppercase transition-all shadow-md hover:scale-[1.02] flex items-center gap-1.5"
                          >
                            <Plus size={18} /> Añadir Pasillo
                          </button>
                        </div>

                        <div className="space-y-3">
                          {aislesData.map((aisle, aisleIndex) => {
                            const titleDraft = aisleBannerDrafts[`t_${aisle.id}`] ?? aisle.bannerTitle ?? "";
                            const subDraft = aisleBannerDrafts[`s_${aisle.id}`] ?? aisle.bannerSubtitle ?? "";
                            const textDirty =
                              titleDraft !== (aisle.bannerTitle ?? "") ||
                              subDraft !== (aisle.bannerSubtitle ?? "");
                            const isAisleOpen = openAislesAdmin[aisle.id] ?? false;

                            const saveBanner = (
                              partial: { bannerImage?: string; bannerTitle?: string; bannerSubtitle?: string },
                              msg?: string,
                            ) => {
                              const newData = {
                                bannerImage: partial.bannerImage ?? aisle.bannerImage ?? "",
                                bannerTitle: partial.bannerTitle ?? aisle.bannerTitle ?? "",
                                bannerSubtitle: partial.bannerSubtitle ?? aisle.bannerSubtitle ?? "",
                              };
                              queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                old
                                  ? { ...old, aisles: old.aisles.map((a: any) => (a.id === aisle.id ? { ...a, ...newData } : a)) }
                                  : old,
                              );
                              if (msg) showToast(msg);
                              updateAisleMut.mutate(
                                { id: aisle.id, data: newData },
                                { onError: refreshMenu },
                              );
                            };

                            return (
                              <div
                                key={aisle.id}
                                className="bg-[#181826] rounded-2xl border border-white/5 overflow-hidden transition-all"
                              >
                                <div className="flex items-center justify-between p-3.5 bg-[#161622]">
                                  <div className="flex items-center gap-2">
                                    <div className="flex flex-col gap-0.5">
                                      <button
                                        onClick={() => moveAisle(aisleIndex, -1)}
                                        disabled={aisleIndex === 0}
                                        className="p-1 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                      >
                                        <ChevronUp size={14} />
                                      </button>
                                      <button
                                        onClick={() => moveAisle(aisleIndex, 1)}
                                        disabled={aisleIndex === aislesData.length - 1}
                                        className="p-1 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                      >
                                        <ChevronDown size={14} />
                                      </button>
                                    </div>
                                    <span className="font-bold text-white text-sm pl-1">{aisle.name}</span>
                                    {aisle.bannerImage && (
                                      <span className="text-[10px] font-bold text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 px-2 py-0.5 rounded-full flex items-center gap-1">
                                        🖼️ Con Banner
                                      </span>
                                    )}
                                  </div>
                                  <div className="flex items-center gap-2">
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setOpenAislesAdmin((p) => ({ ...p, [aisle.id]: !isAisleOpen }))
                                      }
                                      className="px-3 py-1.5 rounded-lg text-xs font-bold text-gray-300 hover:text-white bg-white/5 hover:bg-white/10 transition-colors flex items-center gap-1.5"
                                    >
                                      <Sparkles size={13} className="text-[#ffd025]" />
                                      {isAisleOpen ? "Plegar Banner" : "Editar Banner & IA"}
                                      {isAisleOpen ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                                    </button>
                                    <button
                                      onClick={() => deleteAisleHandler(aisle.id)}
                                      className="text-red-400/60 hover:text-red-400 p-1.5 rounded-lg hover:bg-red-500/10 transition-colors"
                                      title="Eliminar pasillo"
                                    >
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </div>

                                {isAisleOpen && (
                                  <div className="p-4 space-y-3 border-t border-white/5 bg-[#12121e] animate-fade-in">
                                    {aisle.bannerImage && (
                                      <div className="relative rounded-xl overflow-hidden border border-white/10 h-20">
                                        <img
                                          src={aisle.bannerImage}
                                          alt="banner"
                                          className="w-full h-full object-cover"
                                        />
                                        <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-2">
                                          <p className="text-white text-xs font-black truncate">{titleDraft || aisle.name}</p>
                                        </div>
                                      </div>
                                    )}

                                    <div className="flex gap-2">
                                      <input
                                        type="text"
                                        value={aisle.bannerImage || ""}
                                        onChange={(e) => saveBanner({ bannerImage: e.target.value })}
                                        onBlur={(e) =>
                                          resolveImageUrl(e.target.value, (r) =>
                                            saveBanner({ bannerImage: r }, "Banner actualizado"),
                                          )
                                        }
                                        placeholder="URL imagen de fondo para este pasillo..."
                                        className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                      />
                                      <label className="cursor-pointer px-3.5 py-2.5 bg-[#ffd025]/10 text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] rounded-xl text-xs font-black text-center transition flex items-center gap-1.5 border border-[#ffd025]/20">
                                        <Upload size={14} /> Subir
                                        <input
                                          type="file"
                                          accept="image/*"
                                          className="hidden"
                                          onChange={(e) =>
                                            handleImageUpload(
                                              e,
                                              (url) => saveBanner({ bannerImage: url }, "Banner actualizado"),
                                              { skipCompression: true },
                                            )
                                          }
                                        />
                                      </label>
                                      {aisle.bannerImage && (
                                        <button
                                          onClick={() => {
                                            if (aisle.bannerImage?.includes("/storage/objects/")) {
                                              handleDeleteStorageImage(aisle.bannerImage, () =>
                                                saveBanner({ bannerImage: "" }, "Imagen eliminada"),
                                              );
                                            } else {
                                              saveBanner({ bannerImage: "" }, "Imagen eliminada");
                                            }
                                          }}
                                          className="px-3 py-2 bg-red-500/10 text-red-400 hover:bg-red-500 hover:text-white rounded-xl text-xs font-black transition"
                                        >
                                          <Trash2 size={14} />
                                        </button>
                                      )}
                                    </div>

                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                      <input
                                        type="text"
                                        value={titleDraft}
                                        onChange={(e) =>
                                          setAisleBannerDrafts((p) => ({ ...p, [`t_${aisle.id}`]: e.target.value }))
                                        }
                                        placeholder="Título destacado (ej: 🍻 Cervezas Premium #ofertas)"
                                        className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                      />
                                      <input
                                        type="text"
                                        value={subDraft}
                                        onChange={(e) =>
                                          setAisleBannerDrafts((p) => ({ ...p, [`s_${aisle.id}`]: e.target.value }))
                                        }
                                        placeholder="Subtítulo llamativo (ej: 🔥 ¡Aprovecha packs con descuento!)"
                                        className="w-full bg-[#181826] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                      />
                                    </div>

                                    <div className="flex gap-2">
                                      <button
                                        onClick={async () => {
                                          if (!titleDraft.trim() && !subDraft.trim()) {
                                            showToast("Escribe un borrador primero");
                                            return;
                                          }
                                          try {
                                            showToast("✨ Mejorando textos con IA...");
                                            const res = await fetch(
                                              `${import.meta.env.BASE_URL}api/ai/improve-banner`,
                                              {
                                                method: "POST",
                                                headers: { "Content-Type": "application/json" },
                                                body: JSON.stringify({
                                                  title: titleDraft,
                                                  subtitle: subDraft,
                                                  context: aisle.name,
                                                }),
                                              },
                                            );
                                            if (!res.ok) throw new Error("ai failed");
                                            const data = await res.json();
                                            setAisleBannerDrafts((p) => ({
                                              ...p,
                                              [`t_${aisle.id}`]: data.title ?? titleDraft,
                                              [`s_${aisle.id}`]: data.subtitle ?? subDraft,
                                            }));
                                            showToast("✨ Texto mejorado con IA");
                                          } catch {
                                            showToast("Error al mejorar texto");
                                          }
                                        }}
                                        className="flex-1 px-4 py-2.5 bg-gradient-to-r from-[#ff3b8a] to-[#9d4edd] text-white rounded-xl text-xs font-black hover:scale-[1.01] transition flex items-center justify-center gap-1.5 shadow-md shadow-pink-500/20"
                                      >
                                        <Sparkles size={14} /> Mejorar con IA
                                      </button>
                                      {textDirty && (
                                        <button
                                          onClick={() => {
                                            saveBanner(
                                              { bannerTitle: titleDraft, bannerSubtitle: subDraft },
                                              "Textos actualizados",
                                            );
                                            setAisleBannerDrafts((p) => {
                                              const n = { ...p };
                                              delete n[`t_${aisle.id}`];
                                              delete n[`s_${aisle.id}`];
                                              return n;
                                            });
                                          }}
                                          className="flex-1 px-4 py-2.5 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] rounded-xl text-xs font-black hover:scale-[1.01] transition"
                                        >
                                          Guardar Banner
                                        </button>
                                      )}
                                    </div>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>

                        <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-3 border-t border-white/5">
                          <button
                            onClick={() =>
                              reorderAislesMut.mutate(
                                { data: { ids: aislesData.map((a) => a.id) } },
                                {
                                  onSuccess: () => {
                                    setSavedAisles(true);
                                    setTimeout(() => setSavedAisles(false), 2500);
                                  },
                                },
                              )
                            }
                            disabled={reorderAislesMut.isPending}
                            className="w-full sm:w-auto px-6 py-3 rounded-xl font-black uppercase text-xs transition-all flex items-center justify-center gap-2 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] hover:scale-[1.01] disabled:opacity-60 shadow-lg shadow-[#ffd025]/20"
                          >
                            {savedAisles ? (
                              <><CheckCircle size={15} /> Pasillos Guardados</>
                            ) : reorderAislesMut.isPending ? (
                              "Guardando..."
                            ) : (
                              "Guardar Orden de Pasillos"
                            )}
                          </button>

                          {classificationsSubTab === "aisles" && (
                            <button
                              type="button"
                              onClick={() => {
                                setClassificationsSubTab("subcategories");
                                setOpenClassificationsSections((p) => ({ ...p, subcategories: true }));
                              }}
                              className="w-full sm:w-auto px-5 py-3 rounded-xl font-black uppercase text-xs bg-white/5 hover:bg-white/10 text-white border border-white/10 transition-all flex items-center justify-center gap-2"
                            >
                              Siguiente: Subcategorías ➔
                            </button>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* Section 3: Subcategorías & Etiquetas */}
                {(classificationsSubTab === "subcategories" || classificationsSubTab === "all") && (
                  <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 shadow-2xl shadow-black/80 overflow-hidden">
                    <button
                      type="button"
                      onClick={() =>
                        setOpenClassificationsSections((p) => ({ ...p, subcategories: !p.subcategories }))
                      }
                      className="w-full p-6 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors border-b border-white/5"
                    >
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-2xl bg-[#ffd025]/10 border border-[#ffd025]/20 text-[#ffd025] flex items-center justify-center font-black">
                          3
                        </div>
                        <div>
                          <h3 className="text-lg font-black uppercase text-white flex items-center gap-2">
                            <Tag size={18} className="text-[#ffd025]" /> Subcategorías & Etiquetas
                            <span className="text-xs font-bold text-[#ffd025] bg-[#ffd025]/10 px-2 py-0.5 rounded-full border border-[#ffd025]/20">
                              {subcategoriesData.length} etiquetas
                            </span>
                          </h3>
                          <p className="text-xs text-gray-400">Etiquetas opcionales para destacar en las fotos de productos (ej: Pack, Sin Alcohol, 1 Litro).</p>
                        </div>
                      </div>
                      <div className="p-2 rounded-xl bg-white/5 text-gray-400">
                        {openClassificationsSections.subcategories ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                      </div>
                    </button>

                    {openClassificationsSections.subcategories && (
                      <div className="p-6 space-y-5 animate-fade-in">
                        <div className="flex gap-2.5">
                          <input
                            type="text"
                            value={newSubcatName}
                            onChange={(e) => setNewSubcatName(e.target.value)}
                            placeholder="Nueva subcategoría o etiqueta..."
                            onKeyDown={(e) => e.key === "Enter" && handleAddSubcat()}
                            className="flex-1 bg-[#181826] border border-white/10 rounded-xl p-3.5 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                          />
                          <button
                            onClick={handleAddSubcat}
                            className="px-5 py-3.5 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] rounded-xl font-black text-sm uppercase transition-all shadow-md hover:scale-[1.02] flex items-center gap-1.5"
                          >
                            <Plus size={18} /> Añadir
                          </button>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5">
                          {subcategoriesData.map((sub, subIndex) => (
                            <div
                              key={sub.id}
                              className="flex justify-between items-center bg-[#181826] p-3.5 rounded-xl border border-white/5"
                            >
                              <div className="flex items-center gap-1">
                                <div className="flex flex-col gap-0.5">
                                  <button
                                    onClick={() => moveSubcat(subIndex, -1)}
                                    disabled={subIndex === 0}
                                    className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                  >
                                    <ChevronUp size={14} />
                                  </button>
                                  <button
                                    onClick={() => moveSubcat(subIndex, 1)}
                                    disabled={subIndex === subcategoriesData.length - 1}
                                    className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                  >
                                    <ChevronDown size={14} />
                                  </button>
                                </div>
                                <span className="font-bold text-white text-sm pl-1">{sub.name}</span>
                              </div>
                              <button
                                onClick={() => deleteSubcategoryHandler(sub.id)}
                                className="text-red-400/60 hover:text-red-400 p-1.5 rounded-lg hover:bg-red-500/10 transition-colors"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          ))}
                        </div>

                        <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-3 border-t border-white/5">
                          <button
                            onClick={() =>
                              reorderSubcategoriesMut.mutate(
                                { data: { ids: subcategoriesData.map((s) => s.id) } },
                                {
                                  onSuccess: () => {
                                    setSavedSubcats(true);
                                    setTimeout(() => setSavedSubcats(false), 2500);
                                  },
                                },
                              )
                            }
                            disabled={reorderSubcategoriesMut.isPending}
                            className="w-full sm:w-auto px-6 py-3 rounded-xl font-black uppercase text-xs transition-all flex items-center justify-center gap-2 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] hover:scale-[1.01] disabled:opacity-60 shadow-lg shadow-[#ffd025]/20"
                          >
                            {savedSubcats ? (
                              <><CheckCircle size={15} /> Subcategorías Guardadas</>
                            ) : reorderSubcategoriesMut.isPending ? (
                              "Guardando..."
                            ) : (
                              "Guardar Orden de Subcategorías"
                            )}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {adminTab === "social" && <CommunityAdminPanel />}

            {adminTab === "contingency" && (
              <div className="space-y-6 max-w-5xl mx-auto animate-fade-in pb-12">
                {/* Encabezado del Panel de Contingencia */}
                <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-7 rounded-3xl border border-red-500/20 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-2xl shadow-red-950/20">
                  <div>
                    <div className="flex items-center gap-2 text-xs font-bold text-red-400 uppercase tracking-wider mb-1">
                      <AlertTriangle size={16} /> Panel Exclusivo de Autoadministración
                    </div>
                    <h2 className="text-xl sm:text-2xl font-black text-white uppercase flex items-center gap-2.5">
                      <span>Tienda de Contingencia</span>
                      <span className={`text-[10px] px-2.5 py-1 rounded-full font-bold uppercase tracking-wide border ${
                        settingsDraft.contingencyMode
                          ? "bg-red-600/30 text-red-300 border-red-500/50 animate-pulse"
                          : "bg-emerald-500/15 text-emerald-300 border-emerald-500/30"
                      }`}>
                        {settingsDraft.contingencyMode ? "🚨 MODO CONTINGENCIA EN VIVO" : "✅ TIENDA MODO NORMAL"}
                      </span>
                    </h2>
                    <p className="text-xs text-gray-400 mt-1 max-w-2xl">
                      Configura cada pasillo y producto de forma independiente para cuando la tienda opere en modo de contingencia o emergencia por alta demanda.
                    </p>
                  </div>

                  {/* Botón Principal de Guardado */}
                  <div className="flex items-center gap-3 shrink-0">
                    <button
                      type="button"
                      onClick={saveSettings}
                      className="px-5 py-3 bg-gradient-to-r from-red-600 via-amber-500 to-[#ffd025] text-black font-black rounded-xl text-xs uppercase tracking-wider hover:opacity-95 transition-all shadow-lg shadow-red-900/30 flex items-center gap-2"
                    >
                      <CheckCircle size={16} /> Guardar Configuración Contingencia
                    </button>
                  </div>
                </div>

                {/* SECCIÓN 1: Control Maestro y Mensaje Global de Contingencia */}
                <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 space-y-5 shadow-xl">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-white/10">
                    <div>
                      <span className="text-sm font-black text-white uppercase tracking-wider block">
                        1. Estado General del Modo Contingencia
                      </span>
                      <p className="text-xs text-gray-400 mt-0.5">
                        Activa este interruptor cuando desees restringir la tienda pública al catálogo reducido de contingencia y mostrar avisos de emergencia.
                      </p>
                    </div>
                    <label className="flex items-center gap-3 cursor-pointer bg-[#181826] px-4 py-2.5 rounded-2xl border border-white/10 hover:border-white/20 transition-all shrink-0">
                      <span className={`text-xs font-black uppercase ${settingsDraft.contingencyMode ? "text-red-400" : "text-gray-400"}`}>
                        {settingsDraft.contingencyMode ? "Contingencia Activada" : "Contingencia Inactiva"}
                      </span>
                      <input
                        type="checkbox"
                        checked={!!settingsDraft.contingencyMode}
                        onChange={(e) => setSettingsDraft({ ...settingsDraft, contingencyMode: e.target.checked })}
                        className="w-5 h-5 accent-red-500 rounded cursor-pointer"
                      />
                    </label>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-black text-amber-300 uppercase tracking-wider mb-2">
                        Aviso Emergencia al Cliente (Barra Superior Pública)
                      </label>
                      <textarea
                        rows={3}
                        value={settingsDraft.contingencyMessage || ""}
                        onChange={(e) => setSettingsDraft({ ...settingsDraft, contingencyMessage: e.target.value })}
                        placeholder="Ej: 🚨 MODO CONTINGENCIA: Por alta demanda operando con catálogo reducido y entrega express."
                        className="w-full bg-[#12121d] border border-white/10 rounded-xl p-3 text-white text-xs focus:border-amber-400 focus:outline-none placeholder-gray-500"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-black text-amber-300 uppercase tracking-wider mb-2">
                        Banner Especial de Portada para Contingencia
                      </label>
                      <div className="space-y-2">
                        <div className="w-full h-16 rounded-xl border border-white/10 bg-black overflow-hidden relative">
                          <img
                            src={settingsDraft.contingencyBannerImage || "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=1600&auto=format&fit=crop&q=80"}
                            alt="Banner Contingencia"
                            className="w-full h-full object-cover"
                          />
                        </div>
                        <div className="flex gap-2">
                          <input
                            type="text"
                            value={settingsDraft.contingencyBannerImage || ""}
                            onChange={(e) => setSettingsDraft({ ...settingsDraft, contingencyBannerImage: e.target.value })}
                            placeholder="URL del banner de contingencia"
                            className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs"
                          />
                          <label className="bg-amber-500/10 text-amber-300 px-3 py-2 rounded-xl flex items-center cursor-pointer hover:bg-amber-500/20 border border-amber-500/30 text-xs font-bold gap-1 shrink-0">
                            <Upload size={13} />
                            <span>Subir</span>
                            <input
                              type="file"
                              accept="image/*"
                              className="hidden"
                              onChange={(e) => handleImageUpload(e, (url) => setSettingsDraft({ ...settingsDraft, contingencyBannerImage: url }))}
                            />
                          </label>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* SECCIÓN 2: Configuración Individual de Pasillos para Contingencia */}
                <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 space-y-6 shadow-xl">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-white/10">
                    <div>
                      <span className="text-sm font-black text-[#ffd025] uppercase tracking-wider block">
                        2. Configuración Individual por Pasillo y Selección de Productos
                      </span>
                      <p className="text-xs text-gray-400 mt-0.5">
                        Habilita o deshabilita los productos que estarán disponibles durante la contingencia. Usa el buscador rápido abajo para encontrar productos.
                      </p>
                    </div>

                    {/* Acciones globales de pasillos */}
                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                          allStoreAisles.forEach((aisle) => {
                            cfg[aisle] = { ...(cfg[aisle] || {}), enabled: true };
                          });
                          setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                          showToast("Habilitados todos los pasillos en contingencia");
                        }}
                        className="px-3 py-1.5 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 rounded-xl text-[10.5px] font-bold uppercase transition-colors"
                      >
                        ⚡ Habilitar Todos Pasillos
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                          allStoreAisles.forEach((aisle) => {
                            cfg[aisle] = { ...(cfg[aisle] || {}), enabled: false };
                          });
                          setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                          showToast("Deshabilitados todos los pasillos en contingencia");
                        }}
                        className="px-3 py-1.5 bg-red-500/10 hover:bg-red-500/20 text-red-300 border border-red-500/30 rounded-xl text-[10.5px] font-bold uppercase transition-colors"
                      >
                        🚫 Deshabilitar Todos Pasillos
                      </button>
                    </div>
                  </div>

                  {/* BUSCADOR DE PRODUCTOS Y FILTROS EN MODO CONTINGENCIA */}
                  <div className="bg-[#181826] p-4 rounded-2xl border border-amber-500/30 space-y-3">
                    <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
                      <div className="relative flex-1">
                        <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-amber-400" />
                        <input
                          type="text"
                          value={contingencyProductSearch}
                          onChange={(e) => setContingencyProductSearch(e.target.value)}
                          placeholder="🔍 Buscar productos para habilitar/deshabilitar en contingencia..."
                          className="w-full bg-[#12121d] border border-amber-500/30 rounded-xl pl-10 pr-8 py-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                        />
                        {contingencyProductSearch && (
                          <button
                            type="button"
                            onClick={() => setContingencyProductSearch("")}
                            className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-white"
                          >
                            <X size={14} />
                          </button>
                        )}
                      </div>

                      <div className="flex items-center gap-2 shrink-0">
                        <select
                          value={contingencyStatusFilter}
                          onChange={(e) => setContingencyStatusFilter(e.target.value as any)}
                          className="bg-[#12121d] border border-white/10 rounded-xl px-3 py-2.5 text-xs text-white font-bold focus:outline-none cursor-pointer"
                        >
                          <option value="all">Ver Todos los Estados</option>
                          <option value="active">✓ Solo Habilitados</option>
                          <option value="inactive">✕ Solo Inactivos</option>
                        </select>

                        <button
                          type="button"
                          onClick={() => {
                            const filtered = products.filter((p) => {
                              if (contingencyProductSearch.trim()) {
                                const q = normalize(contingencyProductSearch.trim());
                                if (!normalize(p.name).includes(q) && !normalize(p.category || "").includes(q) && !normalize(p.aisle || "").includes(q)) return false;
                              }
                              return true;
                            });
                            batchProductContingency(filtered, true);
                            showToast(`Habilitados ${filtered.length} productos filtrados`);
                          }}
                          className="px-3 py-2 bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-300 border border-emerald-500/40 rounded-xl text-xs font-bold transition-colors shrink-0"
                        >
                          ✓ Marcar Visibles
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const filtered = products.filter((p) => {
                              if (contingencyProductSearch.trim()) {
                                const q = normalize(contingencyProductSearch.trim());
                                if (!normalize(p.name).includes(q) && !normalize(p.category || "").includes(q) && !normalize(p.aisle || "").includes(q)) return false;
                              }
                              return true;
                            });
                            batchProductContingency(filtered, false);
                            showToast(`Deshabilitados ${filtered.length} productos filtrados`);
                          }}
                          className="px-3 py-2 bg-red-500/20 hover:bg-red-500/30 text-red-300 border border-red-500/40 rounded-xl text-xs font-bold transition-colors shrink-0"
                        >
                          ✕ Desmarcar Visibles
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Lista Tarjetas de Pasillos */}
                  <div className="space-y-4">
                    {allStoreAisles.map((aisleName) => {
                      const aisleConfig = (settingsDraft.contingencyAislesConfig || {})[aisleName] || {};
                      const isEnabled = aisleConfig.enabled !== false;
                      const allAisleProds = products.filter((p) => p.aisle === aisleName || p.category === aisleName);
                      
                      const aisleQuery = aisleSearchQueries[aisleName] || "";
                      const filteredAisleProducts = allAisleProds.filter((p) => {
                        if (contingencyStatusFilter === "active" && p.contingencyEnabled === false) return false;
                        if (contingencyStatusFilter === "inactive" && p.contingencyEnabled !== false) return false;
                        if (contingencyProductSearch.trim()) {
                          const q = normalize(contingencyProductSearch.trim());
                          if (!normalize(p.name).includes(q) && !normalize(p.category || "").includes(q) && !normalize(p.aisle || "").includes(q)) return false;
                        }
                        if (aisleQuery.trim()) {
                          const q = normalize(aisleQuery.trim());
                          if (!normalize(p.name).includes(q) && !normalize(p.category || "").includes(q) && !normalize(p.subcategory || "").includes(q)) return false;
                        }
                        return true;
                      });

                      const activeProductsCount = allAisleProds.filter((p) => p.contingencyEnabled !== false).length;

                      if (contingencyProductSearch.trim() && filteredAisleProducts.length === 0) {
                        return null;
                      }

                      return (
                        <div
                          key={aisleName}
                          className={`rounded-2xl border transition-all p-4 sm:p-5 space-y-4 ${
                            isEnabled
                              ? "bg-[#181826] border-white/10 shadow-lg"
                              : "bg-[#12121c]/60 border-white/5 opacity-75"
                          }`}
                        >
                          {/* Encabezado del Pasillo */}
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                            <div className="flex items-center gap-3 min-w-0">
                              <div className={`w-9 h-9 rounded-xl flex items-center justify-center font-black text-sm shrink-0 ${
                                isEnabled ? "bg-amber-500/20 text-amber-300 border border-amber-500/30" : "bg-gray-800 text-gray-500"
                              }`}>
                                <Tag size={16} />
                              </div>
                              <div className="min-w-0">
                                <div className="flex items-center gap-2">
                                  <h3 className="text-sm font-black text-white uppercase tracking-wider truncate">{aisleName}</h3>
                                  <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold ${
                                    isEnabled ? "bg-emerald-500/20 text-emerald-300 border border-emerald-500/30" : "bg-red-500/20 text-red-300 border border-red-500/30"
                                  }`}>
                                    {isEnabled ? "PASILLO ACTIVO EN CONTINGENCIA" : "PASILLO OCULTO EN CONTINGENCIA"}
                                  </span>
                                </div>
                                <p className="text-[11px] text-gray-400 mt-0.5">
                                  {activeProductsCount} de {allAisleProds.length} productos disponibles durante contingencia
                                </p>
                              </div>
                            </div>

                            {/* Controles del Pasillo */}
                            <div className="flex items-center gap-2 shrink-0">
                              <label className="flex items-center gap-2 cursor-pointer bg-[#12121d] px-3.5 py-1.5 rounded-xl border border-white/10 hover:border-white/20 transition-colors">
                                <span className="text-xs font-bold text-gray-300">
                                  {isEnabled ? "Activo" : "Inactivo"}
                                </span>
                                <input
                                  type="checkbox"
                                  checked={isEnabled}
                                  onChange={(e) => {
                                    const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                                    cfg[aisleName] = { ...(cfg[aisleName] || {}), enabled: e.target.checked };
                                    setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                                  }}
                                  className="w-4 h-4 accent-[#ffd025] rounded cursor-pointer"
                                />
                              </label>
                            </div>
                          </div>

                          {/* Ajustes específicos del pasillo si está activo */}
                          {isEnabled && (
                            <div className="space-y-3 pt-1">
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                <div>
                                  <label className="block text-[10.5px] font-bold text-gray-400 uppercase mb-1">
                                    Aviso personalizado para este pasillo en Contingencia (opcional)
                                  </label>
                                  <input
                                    type="text"
                                    value={aisleConfig.noticeText || ""}
                                    onChange={(e) => {
                                      const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                                      cfg[aisleName] = { ...(cfg[aisleName] || {}), noticeText: e.target.value };
                                      setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                                    }}
                                    placeholder="Ej: Stock limitado a marcas principales"
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs focus:border-[#ffd025]"
                                  />
                                </div>

                                <div>
                                  <label className="block text-[10.5px] font-bold text-gray-400 uppercase mb-1">
                                    Banner de pasillo exclusivo para Contingencia (URL / opcional)
                                  </label>
                                  <div className="flex gap-1.5">
                                    <input
                                      type="text"
                                      value={aisleConfig.bannerImage || ""}
                                      onChange={(e) => {
                                        const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                                        cfg[aisleName] = { ...(cfg[aisleName] || {}), bannerImage: e.target.value };
                                        setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                                      }}
                                      placeholder="URL de banner especial"
                                      className="flex-1 bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs"
                                    />
                                    <label className="bg-[#ffd025]/10 text-[#ffd025] px-2.5 py-1.5 rounded-xl flex items-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/20 text-xs font-bold gap-1 shrink-0">
                                      <Upload size={12} />
                                      <input
                                        type="file"
                                        accept="image/*"
                                        className="hidden"
                                        onChange={(e) =>
                                          handleImageUpload(e, (url) => {
                                            const cfg = { ...(settingsDraft.contingencyAislesConfig || {}) };
                                            cfg[aisleName] = { ...(cfg[aisleName] || {}), bannerImage: url };
                                            setSettingsDraft({ ...settingsDraft, contingencyAislesConfig: cfg });
                                          })
                                        }
                                      />
                                    </label>
                                  </div>
                                </div>
                              </div>

                              {/* Barra de búsqueda específica de este pasillo */}
                              <div className="pt-2">
                                <div className="relative">
                                  <Search size={13} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                                  <input
                                    type="text"
                                    value={aisleSearchQueries[aisleName] || ""}
                                    onChange={(e) => setAisleSearchQueries({ ...aisleSearchQueries, [aisleName]: e.target.value })}
                                    placeholder={`🔍 Filtrar en "${aisleName}"...`}
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl pl-8 pr-7 py-2 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                                  />
                                  {aisleSearchQueries[aisleName] && (
                                    <button
                                      type="button"
                                      onClick={() => setAisleSearchQueries({ ...aisleSearchQueries, [aisleName]: "" })}
                                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-white"
                                    >
                                      <X size={12} />
                                    </button>
                                  )}
                                </div>
                              </div>

                              {/* Selección rápida de productos de este pasillo */}
                              <div className="pt-2 border-t border-white/5 space-y-2">
                                <div className="flex items-center justify-between">
                                  <span className="text-[10.5px] font-bold text-gray-400 uppercase tracking-wider">
                                    Productos de "{aisleName}" ({filteredAisleProducts.length} mostrados):
                                  </span>
                                  <div className="flex gap-1.5">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        batchProductContingency(filteredAisleProducts, true);
                                        showToast(`Habilitados todos en "${aisleName}"`);
                                      }}
                                      className="text-[9.5px] px-2.5 py-1 bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 rounded font-bold hover:bg-emerald-500/30 transition-colors"
                                    >
                                      ✓ Marcar Todos
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        batchProductContingency(filteredAisleProducts, false);
                                        showToast(`Desmarcados todos en "${aisleName}"`);
                                      }}
                                      className="text-[9.5px] px-2.5 py-1 bg-red-500/20 text-red-300 border border-red-500/30 rounded font-bold hover:bg-red-500/30 transition-colors"
                                    >
                                      ✕ Desmarcar Todos
                                    </button>
                                  </div>
                                </div>

                                {/* Grid de productos del pasillo */}
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 max-h-60 overflow-y-auto pr-1">
                                  {filteredAisleProducts.map((prod) => {
                                    const prodActive = prod.contingencyEnabled !== false;
                                    return (
                                      <div
                                        key={prod.id}
                                        onClick={() => toggleProductContingency(prod)}
                                        className={`p-2.5 rounded-xl border flex items-center justify-between gap-2 transition-all cursor-pointer select-none ${
                                          prodActive
                                            ? "bg-emerald-950/30 border-emerald-500/50 shadow-md"
                                            : "bg-[#101017]/80 border-white/10 opacity-60 hover:opacity-100"
                                        }`}
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          {prod.image ? (
                                            <img src={prod.image} alt={prod.name} className="w-8 h-8 object-cover rounded-lg bg-black shrink-0 border border-white/10" />
                                          ) : (
                                            <div className="w-8 h-8 bg-white/5 rounded-lg flex items-center justify-center shrink-0">
                                              <Package size={14} className="text-gray-500" />
                                            </div>
                                          )}
                                          <div className="min-w-0">
                                            <p className="text-xs font-bold text-white truncate">{prod.name}</p>
                                            <p className="text-[10px] font-semibold text-[#ffd025]">${prod.price.toLocaleString("es-CL")}</p>
                                          </div>
                                        </div>

                                        <button
                                          type="button"
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            toggleProductContingency(prod);
                                          }}
                                          className={`px-2.5 py-1 rounded-lg text-[10px] font-black uppercase transition-all shrink-0 ${
                                            prodActive
                                              ? "bg-emerald-500 text-black shadow-md"
                                              : "bg-white/10 text-gray-400 border border-white/10 hover:bg-white/20 hover:text-white"
                                          }`}
                                        >
                                          {prodActive ? "✓ Activo" : "Inactivo"}
                                        </button>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Botón Flotante de Guardado Final */}
                <div className="pt-2">
                  <button
                    type="button"
                    onClick={saveSettings}
                    className="w-full py-4 bg-gradient-to-r from-red-600 via-amber-500 to-[#ffd025] text-black rounded-2xl font-black uppercase text-sm tracking-wider hover:scale-[1.002] active:scale-[0.998] transition-all shadow-2xl flex items-center justify-center gap-2"
                  >
                    <CheckCircle size={18} /> Guardar Todos los Ajustes de Contingencia
                  </button>
                </div>
              </div>
            )}

            {adminTab === "products" && (() => {
              const totalCount = products.length;
              const ofertaCount = products.filter((p) => p.oferta).length;
              const contingencyCount = products.filter((p) => p.contingencyEnabled).length;
              const depositoCount = products.filter((p) => p.depositoEnabled).length;

              const specialFilteredProducts = products.filter((p) => {
                if (productFilterSpecial === "oferta") return p.oferta;
                if (productFilterSpecial === "contingency") return p.contingencyEnabled;
                if (productFilterSpecial === "retornable") return p.depositoEnabled;
                return true;
              });

              const isWizardOpen = showAddProductModal || editingProduct !== null;

              return (
                <div className="space-y-6 max-w-7xl mx-auto animate-fade-in">
                  {/* Top Executive KPI Metrics */}
                  <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
                    <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-2xl border border-white/10 shadow-lg shadow-black/40">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[11px] font-bold uppercase tracking-wider text-gray-400">Total Catálogo</span>
                        <div className="p-2 rounded-xl bg-[#ffd025]/10 text-[#ffd025]">
                          <Package size={16} />
                        </div>
                      </div>
                      <div className="text-2xl sm:text-3xl font-black text-white">{totalCount}</div>
                      <p className="text-[10px] text-gray-500 mt-1">Artículos registrados</p>
                    </div>

                    <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-2xl border border-white/10 shadow-lg shadow-black/40">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[11px] font-bold uppercase tracking-wider text-gray-400">En Oferta</span>
                        <div className="p-2 rounded-xl bg-pink-500/10 text-pink-400">
                          <Tag size={16} />
                        </div>
                      </div>
                      <div className="text-2xl sm:text-3xl font-black text-pink-400">{ofertaCount}</div>
                      <p className="text-[10px] text-gray-500 mt-1">Con badge de descuento</p>
                    </div>

                    <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-2xl border border-white/10 shadow-lg shadow-black/40">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[11px] font-bold uppercase tracking-wider text-gray-400">Contingencia</span>
                        <div className="p-2 rounded-xl bg-amber-500/10 text-amber-400">
                          <Ban size={16} />
                        </div>
                      </div>
                      <div className="text-2xl sm:text-3xl font-black text-amber-400">{contingencyCount}</div>
                      <p className="text-[10px] text-gray-500 mt-1">En catálogo de emergencia</p>
                    </div>

                    <div className="bg-[#13131f]/90 backdrop-blur-2xl p-4 sm:p-5 rounded-2xl border border-white/10 shadow-lg shadow-black/40">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[11px] font-bold uppercase tracking-wider text-gray-400">Retornables</span>
                        <div className="p-2 rounded-xl bg-cyan-500/10 text-cyan-400">
                          <CheckCircle size={16} />
                        </div>
                      </div>
                      <div className="text-2xl sm:text-3xl font-black text-cyan-400">{depositoCount}</div>
                      <p className="text-[10px] text-gray-500 mt-1">Con seña de envase</p>
                    </div>
                  </div>

                  {/* Collapsible Step-by-Step Product Wizard (Añadir / Editar) */}
                  {(() => {
                    const renderWizardBlock = (inlineProduct?: Product) => {
                      const activeExtrasCount = [
                        formState.oferta,
                        formState.contingencyEnabled,
                        formState.depositoEnabled,
                        formState.transferenciaEnabled,
                      ].filter(Boolean).length;

                      return (
                        <div className={`bg-[#12121e]/98 backdrop-blur-2xl rounded-3xl border-2 ${inlineProduct ? "border-[#ffd025] ring-4 ring-[#ffd025]/20 my-3" : "border-[#ffd025]/40"} p-4 sm:p-6 shadow-2xl shadow-black/90 space-y-4 animate-fade-in text-left min-w-0 max-w-full overflow-hidden`}>
                          {/* Header con Botón Guardar Principal y Atajo Enter */}
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-white/10">
                            <div>
                              <div className="flex items-center gap-2 text-[11px] font-bold text-[#ffd025] uppercase tracking-wider mb-0.5">
                                <Sparkles size={13} /> Editor de Productos
                                <span className="bg-[#ffd025]/20 text-[#ffd025] border border-[#ffd025]/30 px-2 py-0.5 rounded text-[10px] font-mono">
                                  Enter ↵ = Guardar y Cerrar
                                </span>
                              </div>
                              <h3 className="text-base sm:text-lg font-black text-white uppercase flex items-center gap-2">
                                {inlineProduct
                                  ? `✏️ Editando: ${formState.name || inlineProduct.name}`
                                  : "✨ Añadir Nuevo Producto"}
                              </h3>
                            </div>

                            {/* Acciones Rápidas Superiores: GUARDAR A LA MANO */}
                            <div className="flex items-center gap-2 shrink-0">
                              <button
                                type="button"
                                onClick={saveProduct}
                                className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] font-black uppercase text-xs tracking-wider hover:scale-[1.02] active:scale-[0.98] transition-all shadow-lg shadow-[#ffd025]/25 flex items-center gap-1.5 cursor-pointer ring-2 ring-[#ffd025]/60"
                              >
                                <CheckCircle size={16} />
                                <span>Guardar (Enter ↵)</span>
                              </button>
                              <button
                                type="button"
                                onClick={cancelEditing}
                                className="px-3.5 py-2.5 rounded-xl bg-red-500/10 hover:bg-red-500/20 text-red-300 text-xs font-bold transition-colors border border-red-500/20 flex items-center gap-1 cursor-pointer"
                              >
                                <X size={14} /> Cerrar
                              </button>
                            </div>
                          </div>

                          <form
                            onSubmit={saveProduct}
                            onKeyDown={(e) => {
                              if (e.key === "Enter" && !e.shiftKey) {
                                e.preventDefault();
                                saveProduct(e);
                              }
                            }}
                            className="space-y-4 min-w-0 max-w-full overflow-hidden"
                          >
                            {/* 1. SECCIÓN PRINCIPAL: CARGAR IMAGEN (A LA VISTA Y A LA MANO) */}
                            <div className="rounded-2xl bg-gradient-to-br from-[#1b1b2d] via-[#161626] to-[#12121e] border-2 border-[#ffd025]/50 p-4 sm:p-5 shadow-xl space-y-3 min-w-0 max-w-full overflow-hidden">
                              <div className="flex items-center justify-between">
                                <div className="flex items-center gap-2">
                                  <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-[#ffd025] flex items-center gap-1.5">
                                    <Upload size={16} className="text-[#ffd025]" />
                                    Cargar Imagen del Producto
                                  </span>
                                  <span className={`text-[10px] px-2 py-0.5 rounded font-mono font-bold ${formState.image ? "bg-emerald-500/20 text-emerald-300 border border-emerald-500/30" : "bg-amber-500/20 text-amber-300 border border-amber-500/30"}`}>
                                    {formState.image ? "✓ Foto Asignada" : "⚠️ Sin Foto"}
                                  </span>
                                </div>
                                <span className="text-[11px] text-gray-400 hidden sm:inline">
                                  Acceso rápido e inmediato
                                </span>
                              </div>

                              <div className="flex flex-col sm:flex-row gap-4 items-start min-w-0 max-w-full">
                                {/* Vista previa grande */}
                                <div className="relative w-24 h-24 sm:w-28 sm:h-28 rounded-2xl bg-[#0a0a10] border-2 border-white/15 flex items-center justify-center overflow-hidden shrink-0 shadow-inner group">
                                  {formState.image ? (
                                    <>
                                      <img
                                        src={formState.image}
                                        alt="Preview"
                                        className="w-full h-full object-contain p-1.5 group-hover:scale-105 transition-transform"
                                        onError={(e) => {
                                          (e.currentTarget as HTMLImageElement).src =
                                            "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                        }}
                                      />
                                      <button
                                        type="button"
                                        title="Quitar foto"
                                        onClick={() => {
                                          if (formState.image.includes("/storage/objects/")) {
                                            handleDeleteStorageImage(formState.image, () =>
                                              setFormState({ ...formState, image: "" }),
                                            );
                                          } else {
                                            setFormState({ ...formState, image: "" });
                                          }
                                        }}
                                        className="absolute inset-0 bg-black/75 opacity-0 group-hover:opacity-100 flex flex-col items-center justify-center text-red-400 text-[10px] font-bold transition-opacity cursor-pointer"
                                      >
                                        <Trash2 size={16} />
                                        <span>Quitar</span>
                                      </button>
                                    </>
                                  ) : (
                                    <div className="text-center p-2 text-gray-500 flex flex-col items-center">
                                      <Package size={26} className="text-[#ffd025]/50 mb-1" />
                                      <span className="text-[9px] uppercase font-bold text-gray-400">Sin Foto</span>
                                    </div>
                                  )}
                                </div>

                                {/* Botón Subir Archivo + Input URL */}
                                <div className="flex-1 min-w-0 w-full max-w-full space-y-2.5 overflow-hidden">
                                  <div className="flex flex-wrap gap-2 items-center min-w-0 max-w-full">
                                    <label className="cursor-pointer flex-1 sm:flex-none px-4 py-2.5 bg-[#ffd025] hover:bg-[#ffe066] text-[#0a0a0f] rounded-xl text-xs font-black transition flex items-center justify-center gap-2 shadow-md shadow-[#ffd025]/20 shrink-0">
                                      <Upload size={15} strokeWidth={2.5} />
                                      <span>SUBIR FOTO DESDE TU DISPOSITIVO</span>
                                      <input
                                        type="file"
                                        accept="image/*"
                                        className="hidden"
                                        onChange={(e) =>
                                          handleImageUpload(e, (url) => setFormState({ ...formState, image: url }))
                                        }
                                      />
                                    </label>

                                    <button
                                      type="button"
                                      onClick={() => setMediaPickerOpen(true)}
                                      className="cursor-pointer flex-1 sm:flex-none px-4 py-2.5 bg-white/10 hover:bg-[#ffd025] hover:text-[#0a0a0f] text-gray-200 rounded-xl text-xs font-black transition flex items-center justify-center gap-2 border border-white/15 shadow-sm shrink-0"
                                    >
                                      <FolderOpen size={15} />
                                      <span>ELEGIR DEL ARCHIVO / GALERÍA</span>
                                    </button>

                                    {formState.image && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (formState.image.includes("/storage/objects/")) {
                                            handleDeleteStorageImage(formState.image, () =>
                                              setFormState({ ...formState, image: "" }),
                                            );
                                          } else {
                                            setFormState({ ...formState, image: "" });
                                          }
                                        }}
                                        className="px-3 py-2.5 bg-red-500/10 hover:bg-red-500/20 text-red-400 border border-red-500/20 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shrink-0"
                                      >
                                        <Trash2 size={13} />
                                        <span>Quitar</span>
                                      </button>
                                    )}
                                  </div>

                                  <div className="flex gap-2 min-w-0 max-w-full">
                                    <input
                                      type="text"
                                      placeholder="O escribe / pega el link web de la imagen (https://...)"
                                      value={formState.image.includes("/storage/objects/") ? "" : formState.image}
                                      onChange={(e) => setFormState({ ...formState, image: e.target.value })}
                                      onBlur={(e) =>
                                        resolveImageUrl(e.target.value, (resolved) =>
                                          setFormState((prev) => ({ ...prev, image: resolved })),
                                        )
                                      }
                                      className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                                    />
                                  </div>

                                  {/* Carrusel / Tira rápida de fotos del archivo dentro de la misma tarjeta */}
                                  {quickMediaImages.length > 0 && (
                                    <div className="pt-2.5 border-t border-white/10 space-y-1.5 min-w-0 w-full max-w-full overflow-hidden">
                                      <div className="flex items-center justify-between min-w-0 max-w-full">
                                        <span className="text-[10px] font-bold text-gray-300 uppercase tracking-wider flex items-center gap-1.5 truncate">
                                          <FolderOpen size={12} className="text-[#ffd025] shrink-0" />
                                          <span className="truncate">Fotos de tu Galería ({quickMediaImages.length}) — Clic para asignar directo:</span>
                                        </span>
                                        <button
                                          type="button"
                                          onClick={() => setMediaPickerOpen(true)}
                                          className="text-[10px] text-[#ffd025] hover:text-[#ffe066] font-bold flex items-center gap-1 cursor-pointer transition hover:underline shrink-0 ml-2"
                                        >
                                          <span>Pantalla Completa</span>
                                          <ExternalLink size={10} />
                                        </button>
                                      </div>

                                      <div className="flex items-center gap-2 overflow-x-auto pb-1.5 custom-admin-scrollbar min-w-0 w-full max-w-full scroll-smooth">
                                        {quickMediaImages.slice(0, 25).map((img) => {
                                          const isSelected = formState.image === img.url;
                                          return (
                                            <button
                                              key={img.id}
                                              type="button"
                                              onClick={() => {
                                                setFormState((prev) => ({ ...prev, image: img.url }));
                                                showToast("Foto asignada desde la galería");
                                              }}
                                              title={img.name}
                                              className={`relative w-13 h-13 sm:w-14 sm:h-14 rounded-xl overflow-hidden bg-black/60 border-2 shrink-0 transition-all cursor-pointer group ${
                                                isSelected
                                                  ? "border-[#ffd025] ring-2 ring-[#ffd025]/50 scale-105 shadow-md shadow-[#ffd025]/20"
                                                  : "border-white/15 hover:border-[#ffd025]"
                                              }`}
                                            >
                                              <img
                                                src={img.url}
                                                alt={img.name}
                                                className="w-full h-full object-contain p-1"
                                                loading="lazy"
                                                onError={(e) => {
                                                  (e.currentTarget as HTMLImageElement).src =
                                                    "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80";
                                                }}
                                              />
                                              {isSelected && (
                                                <div className="absolute inset-0 bg-[#ffd025]/30 flex items-center justify-center">
                                                  <div className="w-4 h-4 rounded-full bg-[#ffd025] text-black flex items-center justify-center shadow-md">
                                                    <Check size={11} strokeWidth={3} />
                                                  </div>
                                                </div>
                                              )}
                                            </button>
                                          );
                                        })}

                                        <button
                                          type="button"
                                          onClick={() => setMediaPickerOpen(true)}
                                          className="w-13 h-13 sm:w-14 sm:h-14 rounded-xl border-2 border-dashed border-white/20 hover:border-[#ffd025] text-gray-400 hover:text-[#ffd025] flex flex-col items-center justify-center gap-0.5 shrink-0 transition text-[9px] font-bold uppercase cursor-pointer bg-white/5"
                                          title="Ver todas las fotos o subir nuevas"
                                        >
                                          <FolderOpen size={13} />
                                          <span>Más</span>
                                        </button>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </div>
                            </div>

                            {/* 2. DATOS PRINCIPALES (NOMBRE, PRECIO, CATEGORÍA, PASILLO) */}
                            <div className="rounded-2xl bg-[#181826] border border-white/10 p-4 space-y-3">
                              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                                <div className="md:col-span-2 space-y-1">
                                  <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                    Nombre del Producto *
                                  </label>
                                  <input
                                    required
                                    type="text"
                                    placeholder="Ej: Cerveza Austral Calafate 6x330ml..."
                                    value={formState.name}
                                    onChange={(e) => setFormState({ ...formState, name: e.target.value })}
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
                                  />
                                </div>
                                <div className="space-y-1">
                                  <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                    Precio de Venta ($ CLP) *
                                  </label>
                                  <input
                                    required
                                    type="number"
                                    min="0"
                                    placeholder="Ej: 12990"
                                    value={formState.price}
                                    onChange={(e) => setFormState({ ...formState, price: e.target.value })}
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-[#ffd025] font-black text-xs focus:border-[#ffd025] focus:outline-none"
                                  />
                                </div>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                <div className="space-y-1">
                                  <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                    Categoría Principal *
                                  </label>
                                  <select
                                    required
                                    value={formState.category}
                                    onChange={(e) => setFormState({ ...formState, category: e.target.value })}
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                  >
                                    <option value="">Seleccionar Categoría...</option>
                                    {categories.map((c) => (
                                      <option key={c} value={c}>
                                        {c}
                                      </option>
                                    ))}
                                  </select>
                                </div>

                                <div className="space-y-1">
                                  <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                    Pasillo Separador *
                                  </label>
                                  <select
                                    required
                                    value={formState.aisle}
                                    onChange={(e) => setFormState({ ...formState, aisle: e.target.value })}
                                    className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                  >
                                    <option value="">Seleccionar Pasillo...</option>
                                    {aisles.map((a) => (
                                      <option key={a} value={a}>
                                        {a}
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              </div>
                            </div>

                            {/* 3. OPCIONES SECUNDARIAS (PLEGABLE Y COMPACTO) */}
                            <div className="rounded-2xl bg-[#181826] border border-white/10 overflow-hidden">
                              <button
                                type="button"
                                onClick={() =>
                                  setOpenProductFormSections((p) => ({ ...p, extras: !p.extras }))
                                }
                                className="w-full p-3.5 flex items-center justify-between text-left hover:bg-white/[0.02] transition-colors"
                              >
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-black uppercase text-gray-300">
                                    ⚙️ Opciones Secundarias (Subcategoría, Variantes, Oferta, Contingencia)
                                  </span>
                                  {activeExtrasCount > 0 && (
                                    <span className="text-[10px] px-2 py-0.5 rounded font-mono bg-amber-500/20 text-amber-300">
                                      {activeExtrasCount} activos
                                    </span>
                                  )}
                                </div>
                                <div className="text-gray-400">
                                  {openProductFormSections.extras ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                                </div>
                              </button>

                              {openProductFormSections.extras && (
                                <div className="p-4 pt-0 border-t border-white/5 space-y-3 animate-fade-in mt-3">
                                  {/* Subcategoría */}
                                  <div className="space-y-1">
                                    <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                      Subcategoría / Tag (Opcional)
                                    </label>
                                    <select
                                      value={formState.subcategory}
                                      onChange={(e) => setFormState({ ...formState, subcategory: e.target.value })}
                                      className="w-full bg-[#12121d] border border-white/10 rounded-xl p-2.5 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                    >
                                      <option value="">Sin subcategoría...</option>
                                      {subcategories.map((s) => (
                                        <option key={s} value={s}>
                                          {s}
                                        </option>
                                      ))}
                                    </select>
                                  </div>

                                  {/* Presentaciones */}
                                  <div className="pt-2 border-t border-white/5 space-y-1.5">
                                    <label className="block text-[11px] font-bold uppercase tracking-wider text-gray-400">
                                      Opciones / Presentaciones (Opcional)
                                    </label>
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                      <input
                                        type="text"
                                        placeholder="Título selector (ej: Sabor, Tamaños)"
                                        value={formState.optionsTitle}
                                        onChange={(e) => setFormState({ ...formState, optionsTitle: e.target.value })}
                                        className="bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                      />
                                      <input
                                        type="text"
                                        placeholder="Opciones por coma (ej: Lata 350ml, Botella 1L)"
                                        value={formState.optionsString}
                                        onChange={(e) => setFormState({ ...formState, optionsString: e.target.value })}
                                        className="bg-[#12121d] border border-white/10 rounded-xl p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                      />
                                    </div>
                                  </div>

                                  {/* Toggles */}
                                  <div className="pt-2 border-t border-white/5 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                    {/* Oferta Toggle */}
                                    <div
                                      onClick={() => setFormState({ ...formState, oferta: !formState.oferta })}
                                      className={`p-3 rounded-xl border cursor-pointer transition-all ${
                                        formState.oferta
                                          ? "bg-pink-600/20 border-pink-500 text-white"
                                          : "bg-[#12121d] border-white/10 text-gray-400 hover:border-white/20"
                                      }`}
                                    >
                                      <div className="flex items-center justify-between">
                                        <span className="font-bold text-xs text-white">🏷️ Producto en Oferta</span>
                                        <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[9px] ${formState.oferta ? "bg-pink-500 border-pink-500 text-white" : "border-gray-600"}`}>
                                          {formState.oferta && "✓"}
                                        </span>
                                      </div>
                                    </div>

                                    {/* Contingencia Toggle */}
                                    <div
                                      onClick={() => setFormState({ ...formState, contingencyEnabled: !formState.contingencyEnabled })}
                                      className={`p-3 rounded-xl border cursor-pointer transition-all ${
                                        formState.contingencyEnabled
                                          ? "bg-amber-600/20 border-amber-500 text-white"
                                          : "bg-[#12121d] border-white/10 text-gray-400 hover:border-white/20"
                                      }`}
                                    >
                                      <div className="flex items-center justify-between">
                                        <span className="font-bold text-xs text-white">⚡ Modo Contingencia</span>
                                        <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[9px] ${formState.contingencyEnabled ? "bg-amber-500 border-amber-500 text-white" : "border-gray-600"}`}>
                                          {formState.contingencyEnabled && "✓"}
                                        </span>
                                      </div>
                                    </div>

                                    {/* Retornable Toggle */}
                                    <div className="p-3 rounded-xl bg-[#12121d] border border-white/10 space-y-2">
                                      <div
                                        onClick={() => setFormState({ ...formState, depositoEnabled: !formState.depositoEnabled })}
                                        className="flex items-center justify-between cursor-pointer"
                                      >
                                        <span className="font-bold text-xs text-white">🍾 Envase Retornable</span>
                                        <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[9px] ${formState.depositoEnabled ? "bg-cyan-500 border-cyan-500 text-white" : "border-gray-600"}`}>
                                          {formState.depositoEnabled && "✓"}
                                        </span>
                                      </div>
                                      {formState.depositoEnabled && (
                                        <div className="flex items-center gap-2 pt-1 border-t border-white/5">
                                          <label className="text-[10px] font-bold text-cyan-400">Seña ($):</label>
                                          <input
                                            type="number"
                                            min="0"
                                            value={formState.depositoAmount}
                                            onChange={(e) => setFormState({ ...formState, depositoAmount: e.target.value })}
                                            className="flex-1 bg-[#181826] border border-cyan-500/50 rounded-lg px-2 py-0.5 text-white text-xs"
                                          />
                                        </div>
                                      )}
                                    </div>

                                    {/* Transferencia Fee Toggle */}
                                    <div className="p-3 rounded-xl bg-[#12121d] border border-white/10 space-y-2">
                                      <div
                                        onClick={() => setFormState({ ...formState, transferenciaEnabled: !formState.transferenciaEnabled })}
                                        className="flex items-center justify-between cursor-pointer"
                                      >
                                        <span className="font-bold text-xs text-white">💳 Recargo Transferencia</span>
                                        <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[9px] ${formState.transferenciaEnabled ? "bg-blue-500 border-blue-500 text-white" : "border-gray-600"}`}>
                                          {formState.transferenciaEnabled && "✓"}
                                        </span>
                                      </div>
                                      {formState.transferenciaEnabled && (
                                        <div className="flex items-center gap-2 pt-1 border-t border-white/5">
                                          <label className="text-[10px] font-bold text-blue-400">Monto ($):</label>
                                          <input
                                            type="number"
                                            min="0"
                                            value={formState.transferenciaAmount}
                                            onChange={(e) => setFormState({ ...formState, transferenciaAmount: e.target.value })}
                                            className="flex-1 bg-[#181826] border border-blue-500/50 rounded-lg px-2 py-0.5 text-white text-xs"
                                          />
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              )}
                            </div>

                            {/* Barra de Guardado Inferior */}
                            <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-3 border-t border-white/10">
                              <span className="text-[11px] text-gray-400 font-mono flex items-center gap-1.5">
                                💡 Tip: Puedes presionar <kbd className="px-1.5 py-0.5 bg-white/10 rounded text-[#ffd025] font-bold">Enter ↵</kbd> en cualquier campo para guardar y cerrar.
                              </span>
                              <div className="flex items-center gap-2.5 w-full sm:w-auto justify-end">
                                <button
                                  type="button"
                                  onClick={cancelEditing}
                                  className="px-5 py-2.5 rounded-xl border border-white/10 text-gray-300 hover:text-white hover:bg-white/5 text-xs font-bold transition-colors cursor-pointer"
                                >
                                  Cancelar
                                </button>
                                <button
                                  type="submit"
                                  className="px-7 py-2.5 bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] rounded-xl font-black uppercase text-xs tracking-wider hover:scale-[1.01] active:scale-[0.99] transition-all shadow-lg shadow-[#ffd025]/20 flex items-center justify-center gap-2 cursor-pointer"
                                >
                                  <CheckCircle size={15} />
                                  {inlineProduct ? "Guardar y Cerrar (Enter ↵)" : "Crear Producto"}
                                </button>
                              </div>
                            </div>
                          </form>
                        </div>
                      );
                    };

                    return (
                      <>
                        {/* Modal selector de galería de imágenes para productos */}
                        <MediaLibraryModal
                          isOpen={mediaPickerOpen}
                          onClose={() => setMediaPickerOpen(false)}
                          onSelectImage={(url) => {
                            setFormState((prev) => ({ ...prev, image: url }));
                            showToast("Foto seleccionada desde la galería");
                            loadQuickMediaImages();
                          }}
                          title="Seleccionar Imagen para el Producto"
                        />

                        {/* Top Wizard (solo cuando se añade nuevo producto) */}
                        {showAddProductModal && editingProduct === null && renderWizardBlock()}

                  {/* Main Catalog View Container */}
                  <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-8 rounded-3xl border border-white/10 shadow-2xl shadow-black/80 space-y-6">
                    {/* Catalog Header & Action Tools */}
                    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-white/10">
                      <div className="flex items-center gap-3 flex-wrap">
                        <h2 className="text-xl font-black uppercase text-white flex items-center gap-2">
                          <Package size={22} className="text-[#ffd025]" /> Catálogo ({products.length})
                        </h2>

                        {!isWizardOpen && (
                          <div className="flex items-center gap-2 flex-wrap">
                            <button
                              type="button"
                              onClick={() => {
                                cancelEditing();
                                setShowAddProductModal(true);
                                setProductWizardStep(1);
                              }}
                              className="px-4 py-2 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] rounded-xl font-black text-xs uppercase flex items-center gap-1.5 shadow-md shadow-[#ffd025]/20 hover:scale-[1.02] transition-all"
                            >
                              <Plus size={15} /> + Nuevo Producto
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setAdminTab("settings");
                                setSettingsSubTab("featured");
                                setOpenSettingsSections((p) => ({ ...p, featured: true }));
                              }}
                              className="px-3.5 py-2 bg-[#ffd025]/15 border border-[#ffd025]/30 text-[#ffd025] hover:bg-[#ffd025]/25 rounded-xl font-bold text-xs uppercase flex items-center gap-1.5 transition-all"
                              title="Seleccionar productos de Promociones y Colecciones de Celular"
                            >
                              <Star size={14} /> Destacados & Colecciones
                            </button>
                          </div>
                        )}

                        {/* View Switcher: Grid vs Table */}
                        <div className="flex items-center bg-[#181826] p-1 rounded-xl border border-white/10">
                          <button
                            type="button"
                            onClick={() => setProductViewMode("grid")}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                              productViewMode === "grid"
                                ? "bg-[#ffd025] text-[#141414] shadow"
                                : "text-gray-400 hover:text-white"
                            }`}
                          >
                            <Layers size={13} /> Cuadrícula
                          </button>
                          <button
                            type="button"
                            onClick={() => setProductViewMode("table")}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                              productViewMode === "table"
                                ? "bg-[#ffd025] text-[#141414] shadow"
                                : "text-gray-400 hover:text-white"
                            }`}
                          >
                            <FileSpreadsheet size={13} /> Lista
                          </button>
                        </div>
                      </div>

                      {/* Excel & Bulk Controls (Icon-only, single line) */}
                      <div className="flex items-center gap-2 flex-nowrap shrink-0">
                        {/* Respaldo Completo ZIP con fotos y catálogo */}
                        <button
                          type="button"
                          onClick={handleExportBackupZip}
                          title="Descargar Copia de Seguridad Completa con imágenes (.ZIP) para no perder nada al actualizar en Render o GitHub"
                          className="flex items-center justify-center p-2.5 bg-purple-600/20 border border-purple-500/30 text-purple-300 rounded-xl hover:bg-purple-600/35 transition-colors shrink-0"
                        >
                          <Archive size={15} />
                        </button>

                        {/* Restaurar Respaldo ZIP o JSON */}
                        <label
                          title="Restaurar Copia de Seguridad (.ZIP con imágenes o .JSON) sin perder nada"
                          className="flex items-center justify-center p-2.5 bg-emerald-600/20 border border-emerald-500/30 text-emerald-400 rounded-xl hover:bg-emerald-600/35 transition-colors cursor-pointer shrink-0"
                        >
                          <ShieldCheck size={15} />
                          <input
                            type="file"
                            accept=".zip, .json"
                            className="hidden"
                            onChange={handleRestoreBackup}
                          />
                        </label>

                        <button
                          type="button"
                          onClick={handleDownloadProductsExcelWithImages}
                          title="Descargar Excel completo con imágenes incrustadas en máxima calidad (nombre, pasillo, categoría, subcategoría, precio, imagen)"
                          className="flex items-center gap-1.5 px-3 py-2 bg-green-700/20 border border-green-600/40 text-green-400 rounded-xl hover:bg-green-700/35 transition-colors shrink-0 font-bold text-xs"
                        >
                          <Download size={15} />
                          <span className="hidden sm:inline">Excel con Fotos</span>
                        </button>
                        <button
                          type="button"
                          onClick={handleDownloadTemplate}
                          title="Descargar plantilla Excel vacía con formato"
                          className="flex items-center justify-center p-2.5 bg-amber-500/20 border border-amber-500/30 text-amber-400 rounded-xl hover:bg-amber-500/30 transition-colors shrink-0"
                        >
                          <FileSpreadsheet size={15} />
                        </button>
                        <label
                          title="Subir planilla Excel para importar productos"
                          className="flex items-center justify-center p-2.5 bg-blue-600/20 border border-blue-500/30 text-blue-400 rounded-xl hover:bg-blue-600/30 transition-colors cursor-pointer shrink-0"
                        >
                          <Upload size={15} />
                          <input
                            type="file"
                            accept=".xlsx, .xls, .csv"
                            className="hidden"
                            onChange={handleImportExcel}
                          />
                        </label>
                        <button
                          type="button"
                          onClick={handleDeleteAllProducts}
                          title="Eliminar todos los productos"
                          className="flex items-center justify-center p-2.5 bg-red-600/20 border border-red-500/30 text-red-400 rounded-xl hover:bg-red-600/30 transition-colors shrink-0"
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
                    </div>

                    {/* Filter & Search Bar */}
                    <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
                      {/* Special Filters */}
                      <div className="flex items-center gap-2 overflow-x-auto pb-1 scrollbar-hide">
                        {[
                          { id: "all", label: "Todos", count: totalCount },
                          { id: "oferta", label: "🔥 Ofertas", count: ofertaCount },
                          { id: "contingency", label: "⚡ Contingencia", count: contingencyCount },
                          { id: "retornable", label: "🔄 Retornables", count: depositoCount },
                        ].map((fil) => (
                          <button
                            key={fil.id}
                            type="button"
                            onClick={() => setProductFilterSpecial(fil.id as any)}
                            className={`px-3 py-1.5 rounded-full text-xs font-bold whitespace-nowrap transition-all flex items-center gap-1.5 border ${
                              productFilterSpecial === fil.id
                                ? "bg-[#ffd025] text-[#141414] border-[#ffd025] shadow-md shadow-[#ffd025]/20"
                                : "bg-[#181826] text-gray-400 hover:text-white border-white/10"
                            }`}
                          >
                            <span>{fil.label}</span>
                            <span
                              className={`text-[10px] px-1.5 py-0.2 rounded-full font-black ${
                                productFilterSpecial === fil.id ? "bg-black/20 text-[#141414]" : "bg-white/10 text-gray-300"
                              }`}
                            >
                              {fil.count}
                            </span>
                          </button>
                        ))}
                      </div>

                      {/* Search Input */}
                      <div className="relative flex-1 max-w-sm">
                        <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none">
                          🔍
                        </span>
                        <input
                          type="text"
                          placeholder="Buscar producto por nombre, pasillo..."
                          value={adminProductSearch}
                          onChange={(e) => setAdminProductSearch(e.target.value)}
                          className="w-full bg-[#181826] border border-white/10 rounded-xl pl-9 pr-9 py-2.5 text-white text-xs sm:text-sm focus:border-[#ffd025] focus:outline-none placeholder:text-gray-500"
                        />
                        {adminProductSearch && (
                          <button
                            type="button"
                            onClick={() => setAdminProductSearch("")}
                            className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-white transition-colors"
                          >
                            ✕
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Category Scroll Tabs */}
                    <div
                      ref={adminCategoriesRef}
                      className={`flex gap-2 overflow-x-auto scrollbar-hide pb-2 border-b border-white/10 scroll-fade-x scroll-smooth select-none transition-opacity ${
                        adminProductSearch ? "opacity-30 pointer-events-none" : ""
                      }`}
                    >
                      {["Todas", ...categories].map((cat) => {
                        const count =
                          cat === "Todas"
                            ? specialFilteredProducts.length
                            : specialFilteredProducts.filter((p) => p.category === cat).length;
                        const isActive = adminCategory === cat;
                        return (
                          <button
                            key={cat}
                            type="button"
                            onClick={() => setAdminCategory(cat)}
                            className={`flex items-center gap-2 px-3.5 py-2 rounded-xl whitespace-nowrap text-xs font-bold uppercase transition-colors shrink-0 border ${
                              isActive
                                ? "bg-[#ffd025] text-[#141414] border-[#ffd025] shadow-md shadow-[#ffd025]/20"
                                : "bg-[#181826] text-gray-300 border-white/5 hover:border-white/20 hover:text-white"
                            }`}
                          >
                            <span>{cat}</span>
                            <span
                              className={`text-[10px] px-1.5 py-0.5 rounded-full font-bold ${
                                isActive ? "bg-[#141414]/20 text-[#141414]" : "bg-white/10 text-gray-400"
                              }`}
                            >
                              {count}
                            </span>
                          </button>
                        );
                      })}
                    </div>

                    {/* Product Listing Grouped by Aisles with Collapsible Accordions */}
                    <div className="space-y-4">
                      {(() => {
                        const normalizeStr = (s: string) =>
                          s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
                        const filteredProducts = (() => {
                          const q = normalizeStr(adminProductSearch.trim());
                          if (q.length > 0) {
                            const words = q.split(/\s+/).filter(Boolean);
                            return specialFilteredProducts.filter((p) => {
                              const haystack = normalizeStr(
                                [p.name, p.category, p.aisle, p.subcategory ?? ""].join(" "),
                              );
                              return words.every((w) => haystack.includes(w));
                            });
                          }
                          return adminCategory === "Todas"
                            ? specialFilteredProducts
                            : specialFilteredProducts.filter((p) => p.category === adminCategory);
                        })();

                        const groups = aisles.map((a) => ({
                          aisle: a,
                          items: filteredProducts.filter((p) => p.aisle === a),
                        }));
                        const ungrouped = filteredProducts.filter((p) => !aisles.includes(p.aisle));
                        if (ungrouped.length > 0) {
                          groups.push({ aisle: "Sin pasillo", items: ungrouped });
                        }

                        const activeGroups = groups.filter((g) => g.items.length > 0);

                        if (activeGroups.length === 0) {
                          return (
                            <div className="text-center py-12 bg-white/[0.02] rounded-2xl border border-white/5">
                              <Package size={36} className="mx-auto mb-2 text-[#ffd025]/30" />
                              <p className="text-sm font-bold text-gray-400 uppercase">
                                No se encontraron productos con los filtros seleccionados
                              </p>
                            </div>
                          );
                        }

                        return activeGroups.map(({ aisle, items }) => {
                          const isCollapsed = collapsedAisles[aisle];
                          const groupIds = items.map((p) => p.id);

                          return (
                            <div
                              key={aisle}
                              className="rounded-2xl border border-white/10 overflow-hidden bg-[#161624] shadow-md"
                            >
                              {/* Aisle Accordion Header */}
                              <button
                                type="button"
                                onClick={() =>
                                  setCollapsedAisles({ ...collapsedAisles, [aisle]: !isCollapsed })
                                }
                                className="w-full flex items-center justify-between p-4 bg-[#191929] hover:bg-[#1f1f33] transition-colors border-b border-white/5"
                              >
                                <div className="flex items-center gap-2.5">
                                  <div className="p-1 rounded-lg bg-white/5 text-[#ffd025]">
                                    {isCollapsed ? <ChevronRight size={16} /> : <ChevronDown size={16} />}
                                  </div>
                                  <span className="font-black uppercase text-sm text-white tracking-wide">
                                    {aisle}
                                  </span>
                                  <span className="text-[10px] bg-[#ffd025]/15 text-[#ffd025] px-2.5 py-0.5 rounded-full font-black border border-[#ffd025]/20">
                                    {items.length} productos
                                  </span>
                                </div>
                                <span className="text-[11px] text-gray-400 font-medium">
                                  {isCollapsed ? "Clic para desplegar" : "Clic para plegar"}
                                </span>
                              </button>

                              {!isCollapsed && (
                                <div className="p-4 sm:p-5">
                                  {/* Grid Mode */}
                                  {productViewMode === "grid" ? (
                                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                                      {items.map((product) => (
                                        <Fragment key={product.id}>
                                          <div
                                            id={`product-card-${product.id}`}
                                            className={`group relative rounded-2xl border p-4 transition-all flex flex-col justify-between ${
                                              editingProduct === product.id
                                                ? "bg-[#181828] border-[#ffd025] ring-2 ring-[#ffd025]/50 shadow-2xl shadow-[#ffd025]/10"
                                                : product.hidden
                                                ? "bg-[#14141e]/50 border-white/5 opacity-50"
                                                : "bg-[#181828] border-white/10 hover:border-[#ffd025]/50 hover:shadow-xl hover:shadow-black/60"
                                            }`}
                                          >
                                            {/* Card Top: Badges & Photo */}
                                            <div>
                                              <div className="relative w-full h-36 rounded-xl bg-[#12121d] border border-white/5 mb-3 flex items-center justify-center overflow-hidden">
                                                {product.image ? (
                                                  <img
                                                    loading="lazy"
                                                    src={product.image}
                                                    alt={product.name}
                                                    className="w-full h-full object-contain p-2 group-hover:scale-105 transition-transform duration-300"
                                                    onError={(e) => {
                                                      (e.currentTarget as HTMLImageElement).style.display = "none";
                                                    }}
                                                  />
                                                ) : (
                                                  <div className="text-gray-600 flex flex-col items-center">
                                                    <Package size={30} className="mb-1 opacity-40" />
                                                    <span className="text-[9px] uppercase font-bold">Sin foto</span>
                                                  </div>
                                                )}

                                                {/* Badges on image */}
                                                <div className="absolute top-2 left-2 flex flex-col gap-1">
                                                  {product.oferta && (
                                                    <span className="text-[9px] font-black uppercase text-white bg-pink-600 px-2 py-0.5 rounded-md shadow-md">
                                                      🔥 Oferta
                                                    </span>
                                                  )}
                                                  {product.contingencyEnabled && (
                                                    <span className="text-[9px] font-black uppercase text-white bg-amber-600 px-2 py-0.5 rounded-md shadow-md">
                                                      ⚡ Contingencia
                                                    </span>
                                                  )}
                                                  {product.depositoEnabled && (
                                                    <span className="text-[9px] font-black uppercase text-white bg-cyan-600 px-2 py-0.5 rounded-md shadow-md">
                                                      🍾 Retornable
                                                    </span>
                                                  )}
                                                </div>

                                                {/* Visibility badge */}
                                                {product.hidden && (
                                                  <div className="absolute inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center">
                                                    <span className="text-xs font-bold text-gray-300 flex items-center gap-1">
                                                      <EyeOff size={14} /> Oculto
                                                    </span>
                                                  </div>
                                                )}
                                              </div>

                                              {/* Details */}
                                              <h4 className="font-bold text-white text-sm line-clamp-2 leading-snug mb-1">
                                                {product.name}
                                              </h4>

                                              <div className="flex flex-wrap items-center gap-1.5 text-[10px] text-gray-400 mb-2">
                                                <span className="bg-white/5 px-2 py-0.5 rounded-md border border-white/5">
                                                  {product.category}
                                                </span>
                                                {product.subcategory && (
                                                  <span className="bg-[#ffd025]/10 text-[#ffd025] px-2 py-0.5 rounded-md border border-[#ffd025]/20">
                                                    {product.subcategory}
                                                  </span>
                                                )}
                                              </div>

                                              {product.options && product.options.length > 0 && (
                                                <p className="text-[10px] text-blue-400 line-clamp-1 mb-2">
                                                  Opciones: {product.options.join(", ")}
                                                </p>
                                              )}
                                            </div>

                                            {/* Card Bottom: Price and Actions */}
                                            <div className="pt-3 border-t border-white/5 space-y-3">
                                              <div className="flex items-center justify-between">
                                                <span className="text-lg font-black text-[#ffd025]">
                                                  ${product.price.toLocaleString("es-CL")}
                                                </span>
                                                <div className="flex items-center gap-1">
                                                  <button
                                                    title="Subir posición"
                                                    onClick={() => moveProduct(product.id, "up", groupIds)}
                                                    disabled={groupIds.indexOf(product.id) === 0}
                                                    className="p-1 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 transition-colors"
                                                  >
                                                    <ChevronUp size={15} />
                                                  </button>
                                                  <button
                                                    title="Bajar posición"
                                                    onClick={() => moveProduct(product.id, "down", groupIds)}
                                                    disabled={groupIds.indexOf(product.id) === groupIds.length - 1}
                                                    className="p-1 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 transition-colors"
                                                  >
                                                    <ChevronDown size={15} />
                                                  </button>
                                                </div>
                                              </div>

                                              <div className="flex items-center justify-between gap-1 pt-1">
                                                {/* Social Share toggle */}
                                                <button
                                                  title={
                                                    product.publishedSocial
                                                      ? "Publicado en redes ✓"
                                                      : "Marcar publicado en redes"
                                                  }
                                                  onClick={() => {
                                                    const newVal = !product.publishedSocial;
                                                    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                                      old
                                                        ? {
                                                            ...old,
                                                            products: old.products.map((p: any) =>
                                                              p.id === product.id ? { ...p, publishedSocial: newVal } : p,
                                                            ),
                                                          }
                                                        : old,
                                                    );
                                                    updateProductMut.mutate({
                                                      id: product.id,
                                                      data: { ...product, publishedSocial: newVal },
                                                    });
                                                  }}
                                                  className={`p-2 rounded-xl text-xs transition-colors ${
                                                    product.publishedSocial
                                                      ? "bg-green-500/20 text-green-400"
                                                      : "bg-white/5 text-gray-500 hover:text-green-400"
                                                  }`}
                                                >
                                                  <CheckCircle size={15} />
                                                </button>

                                                {/* Hide/Show Toggle */}
                                                <button
                                                  title={product.hidden ? "Mostrar en tienda" : "Ocultar de la tienda"}
                                                  onClick={() => {
                                                    const newVal = !product.hidden;
                                                    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                                      old
                                                        ? {
                                                            ...old,
                                                            products: old.products.map((p: any) =>
                                                              p.id === product.id ? { ...p, hidden: newVal } : p,
                                                            ),
                                                          }
                                                        : old,
                                                    );
                                                    updateProductMut.mutate({
                                                      id: product.id,
                                                      data: { ...product, hidden: newVal },
                                                    });
                                                  }}
                                                  className={`p-2 rounded-xl text-xs transition-colors ${
                                                    product.hidden
                                                      ? "bg-amber-500/20 text-amber-400"
                                                      : "bg-white/5 text-gray-500 hover:text-amber-400"
                                                  }`}
                                                >
                                                  {product.hidden ? <EyeOff size={15} /> : <Eye size={15} />}
                                                </button>

                                                {/* Edit Button */}
                                                <button
                                                  type="button"
                                                  onClick={() => {
                                                    if (editingProduct === product.id) {
                                                      cancelEditing();
                                                    } else {
                                                      startEditing(product);
                                                      setTimeout(() => {
                                                        const el = document.getElementById(`product-editor-${product.id}`);
                                                        if (el) {
                                                          el.scrollIntoView({ behavior: "smooth", block: "nearest" });
                                                        }
                                                      }, 50);
                                                    }
                                                  }}
                                                  className={`flex-1 py-1.5 px-3 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1 cursor-pointer ${
                                                    editingProduct === product.id
                                                      ? "bg-[#ffd025] text-[#141414] hover:bg-[#e5b81a] shadow-md shadow-[#ffd025]/20 font-black"
                                                      : "bg-blue-500/10 hover:bg-blue-500/20 text-blue-400"
                                                  }`}
                                                >
                                                  <Pencil size={13} /> {editingProduct === product.id ? "Plegar" : "Editar"}
                                                </button>

                                                {/* Delete Button */}
                                                <button
                                                  onClick={() => deleteProductHandler(product.id)}
                                                  className="p-2 rounded-xl bg-red-500/10 hover:bg-red-500/20 text-red-400 text-xs transition-colors"
                                                  title="Eliminar producto"
                                                >
                                                  <Trash2 size={15} />
                                                </button>
                                              </div>
                                            </div>
                                          </div>

                                          {/* Editor de producto inline desplegable a la misma altura */}
                                          {editingProduct === product.id && (
                                            <div
                                              id={`product-editor-${product.id}`}
                                              className="col-span-1 sm:col-span-2 lg:col-span-3 xl:col-span-4 w-full min-w-0 max-w-full overflow-hidden"
                                            >
                                              {renderWizardBlock(product)}
                                            </div>
                                          )}
                                        </Fragment>
                                      ))}
                                    </div>
                                  ) : (
                                    /* Table Mode */
                                    <div className="overflow-x-auto">
                                      <table className="w-full text-left text-xs">
                                        <thead>
                                          <tr className="border-b border-white/10 text-gray-400 uppercase text-[10px]">
                                            <th className="py-2.5 px-3">Producto</th>
                                            <th className="py-2.5 px-3">Categoría</th>
                                            <th className="py-2.5 px-3">Precio</th>
                                            <th className="py-2.5 px-3">Atributos</th>
                                            <th className="py-2.5 px-3 text-right">Acciones</th>
                                          </tr>
                                        </thead>
                                        <tbody className="divide-y divide-white/5">
                                          {items.map((product) => (
                                            <Fragment key={product.id}>
                                              <tr
                                                id={`product-row-${product.id}`}
                                                className={`transition-colors group ${
                                                  editingProduct === product.id
                                                    ? "bg-[#ffd025]/10 border-l-4 border-l-[#ffd025]"
                                                    : "hover:bg-white/[0.02]"
                                                }`}
                                              >
                                              <td className="py-3 px-3">
                                                <div className="flex items-center gap-3">
                                                  <div className="w-10 h-10 rounded-lg bg-[#12121d] border border-white/5 flex items-center justify-center overflow-hidden shrink-0">
                                                    {product.image ? (
                                                      <img
                                                        src={product.image}
                                                        alt=""
                                                        className="w-full h-full object-contain p-0.5"
                                                      />
                                                    ) : (
                                                      <Package size={16} className="text-gray-600" />
                                                    )}
                                                  </div>
                                                  <div>
                                                    <p className="font-bold text-white text-xs sm:text-sm">{product.name}</p>
                                                    {product.options && product.options.length > 0 && (
                                                      <p className="text-[10px] text-blue-400">{product.options.join(", ")}</p>
                                                    )}
                                                  </div>
                                                </div>
                                              </td>
                                              <td className="py-3 px-3 text-gray-300">
                                                <span>{product.category}</span>
                                                {product.subcategory && (
                                                  <span className="ml-1 text-[10px] text-[#ffd025]">({product.subcategory})</span>
                                                )}
                                              </td>
                                              <td className="py-3 px-3 font-black text-[#ffd025] text-sm whitespace-nowrap">
                                                ${product.price.toLocaleString("es-CL")}
                                              </td>
                                              <td className="py-3 px-3">
                                                <div className="flex items-center gap-1 flex-wrap">
                                                  {product.oferta && (
                                                    <span className="text-[9px] bg-pink-500/20 text-pink-400 px-1.5 py-0.5 rounded font-bold">
                                                      Oferta
                                                    </span>
                                                  )}
                                                  {product.contingencyEnabled && (
                                                    <span className="text-[9px] bg-amber-500/20 text-amber-400 px-1.5 py-0.5 rounded font-bold">
                                                      Contingencia
                                                    </span>
                                                  )}
                                                  {product.depositoEnabled && (
                                                    <span className="text-[9px] bg-cyan-500/20 text-cyan-400 px-1.5 py-0.5 rounded font-bold">
                                                      Retornable
                                                    </span>
                                                  )}
                                                  {product.hidden && (
                                                    <span className="text-[9px] bg-gray-700 text-gray-300 px-1.5 py-0.5 rounded font-bold">
                                                      Oculto
                                                    </span>
                                                  )}
                                                </div>
                                              </td>
                                              <td className="py-3 px-3 text-right">
                                                <div className="flex items-center justify-end gap-1.5">
                                                  <button
                                                    type="button"
                                                    onClick={() => {
                                                      if (editingProduct === product.id) {
                                                        cancelEditing();
                                                      } else {
                                                        startEditing(product);
                                                        setTimeout(() => {
                                                          const el = document.getElementById(`product-editor-table-${product.id}`);
                                                          if (el) {
                                                            el.scrollIntoView({ behavior: "smooth", block: "nearest" });
                                                          }
                                                        }, 50);
                                                      }
                                                    }}
                                                    className={`p-1.5 rounded-lg transition-colors cursor-pointer ${
                                                      editingProduct === product.id
                                                        ? "bg-[#ffd025] text-black shadow font-bold"
                                                        : "text-blue-400 hover:bg-blue-400/10"
                                                    }`}
                                                    title={editingProduct === product.id ? "Plegar editor" : "Editar"}
                                                  >
                                                    <Pencil size={15} />
                                                  </button>
                                                  <button
                                                    onClick={() => deleteProductHandler(product.id)}
                                                    className="p-1.5 text-red-400 hover:bg-red-400/10 rounded-lg transition-colors"
                                                    title="Eliminar"
                                                  >
                                                    <Trash2 size={15} />
                                                  </button>
                                                </div>
                                              </td>
                                            </tr>
                                            {editingProduct === product.id && (
                                              <tr id={`product-editor-table-${product.id}`}>
                                                <td colSpan={5} className="p-3 sm:p-5 bg-[#0e0e18] border-y-2 border-[#ffd025] min-w-0 max-w-full overflow-hidden">
                                                  {renderWizardBlock(product)}
                                                </td>
                                              </tr>
                                            )}
                                          </Fragment>
                                        ))}
                                      </tbody>
                                      </table>
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        });
                      })()}
                    </div>
                  </div>
                  </>
                );
              })()}
            </div>
          );
        })()}

            </main>
          </div>
        </div>
      )}

      <style
        dangerouslySetInnerHTML={{
          __html: `
        .scrollbar-hide::-webkit-scrollbar { display: none; }
        .scrollbar-hide { -ms-overflow-style: none; scrollbar-width: none; }
        .scroll-fade-x {
          -webkit-mask-image: linear-gradient(to right, transparent 0, black 28px, black calc(100% - 28px), transparent 100%);
          mask-image: linear-gradient(to right, transparent 0, black 28px, black calc(100% - 28px), transparent 100%);
        }
        @keyframes slideInRight { from { transform: translateX(100%); } to { transform: translateX(0); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .animate-slide-in-right { animation: slideInRight 0.3s cubic-bezier(0.16, 1, 0.3, 1); }
        .animate-fade-in { animation: fadeIn 0.4s ease-out forwards; }
      `,
        }}
      />
    </div>
  );
}

const ORDER_STATUSES = [
  {
    value: "Confirmado y en Preparación",
    label: "Confirmado y en Preparación",
    color: "bg-blue-500/15 text-blue-300 border-blue-500/40",
    activeColor: "bg-blue-500 text-white border-blue-500",
    whatsapp: true,
  },
  {
    value: "Listo para Retirar",
    label: "Listo para Retirar",
    color: "bg-amber-500/15 text-amber-300 border-amber-500/40",
    activeColor: "bg-amber-500 text-[#141414] border-amber-500",
    whatsapp: true,
  },
  {
    value: "Delivery en Camino",
    label: "Delivery en Camino",
    color: "bg-emerald-500/15 text-emerald-300 border-emerald-500/40",
    activeColor: "bg-emerald-500 text-[#141414] border-emerald-500",
    whatsapp: true,
  },
  {
    value: "Entregado",
    label: "✓ Entregado",
    color: "bg-green-700/20 text-green-400 border-green-600/40",
    activeColor: "bg-green-600 text-white border-green-600",
    whatsapp: false,
  },
] as const;

function DeliveryLocationsAdminPanel() {
  const queryClient = useQueryClient();
  const { data: locations = [] } = useListDeliveryLocations();
  const createMut = useCreateDeliveryLocation();
  const updateMut = useUpdateDeliveryLocation();
  const deleteMut = useDeleteDeliveryLocation();

  const [form, setForm] = useState({ name: "", price: "" });
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editForm, setEditForm] = useState({ name: "", price: "" });

  const invalidate = () => queryClient.invalidateQueries({ queryKey: getListDeliveryLocationsQueryKey() });

  const handleCreate = () => {
    const price = parseInt(form.price, 10);
    if (!form.name.trim() || isNaN(price)) return;
    createMut.mutate({ data: { name: form.name.trim(), price } }, { onSuccess: () => { setForm({ name: "", price: "" }); invalidate(); } });
  };

  const startEdit = (loc: DeliveryLocation) => {
    setEditingId(loc.id);
    setEditForm({ name: loc.name, price: String(loc.price) });
  };

  const handleUpdate = (id: number) => {
    const price = parseInt(editForm.price, 10);
    if (!editForm.name.trim() || isNaN(price)) return;
    updateMut.mutate({ id, data: { name: editForm.name.trim(), price } }, { onSuccess: () => { setEditingId(null); invalidate(); } });
  };

  const toggleActive = (loc: DeliveryLocation) => {
    updateMut.mutate({ id: loc.id, data: { active: !loc.active } }, { onSuccess: invalidate });
  };

  const handleDelete = (id: number) => {
    deleteMut.mutate({ id }, { onSuccess: invalidate });
  };

  return (
    <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20 max-w-2xl animate-fade-in">
      <h2 className="text-xl font-black uppercase mb-6 flex items-center gap-2">
        <span className="text-[#ffd025]">📍</span> Zonas de Delivery
      </h2>
      <div className="space-y-3 mb-6">
        {locations.length === 0 && (
          <p className="text-gray-500 text-sm">No hay zonas configuradas. Agrega una abajo.</p>
        )}
        {locations.map(loc => (
          <div key={loc.id} className="flex items-center gap-3 bg-[#141414] rounded-xl p-3 border border-gray-800">
            {editingId === loc.id ? (
              <>
                <input
                  value={editForm.name}
                  onChange={e => setEditForm({ ...editForm, name: e.target.value })}
                  className="flex-1 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-lg px-3 py-1.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                  placeholder="Nombre"
                />
                <input
                  type="number"
                  value={editForm.price}
                  onChange={e => setEditForm({ ...editForm, price: e.target.value })}
                  className="w-28 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-lg px-3 py-1.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                  placeholder="Precio"
                />
                <button onClick={() => handleUpdate(loc.id)} className="text-[#ffd025] hover:text-[#e5b81a] font-bold text-xs px-3 py-1.5 bg-[#ffd025]/10 rounded-lg">Guardar</button>
                <button onClick={() => setEditingId(null)} className="text-gray-500 hover:text-gray-300 text-xs px-2 py-1.5">✕</button>
              </>
            ) : (
              <>
                <span className={`w-2 h-2 rounded-full flex-shrink-0 ${loc.active ? "bg-green-400" : "bg-gray-600"}`} />
                <span className="flex-1 text-sm font-semibold text-white">{loc.name}</span>
                <span className="text-[#ffd025] font-black text-sm">${loc.price.toLocaleString("es-CL")}</span>
                <button
                  onClick={() => toggleActive(loc)}
                  className={`text-xs px-2.5 py-1 rounded-lg font-bold transition-all ${loc.active ? "bg-green-900/40 text-green-400 hover:bg-green-900/60" : "bg-gray-800 text-gray-500 hover:text-gray-300"}`}
                >
                  {loc.active ? "Activo" : "Inactivo"}
                </button>
                <button onClick={() => startEdit(loc)} className="text-gray-400 hover:text-[#ffd025] transition-colors">
                  <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/></svg>
                </button>
                <button onClick={() => handleDelete(loc.id)} className="text-gray-600 hover:text-red-400 transition-colors">
                  <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                </button>
              </>
            )}
          </div>
        ))}
      </div>
      <div className="border-t border-gray-800 pt-5 space-y-3">
        <p className="text-xs font-bold text-gray-400 uppercase">Nueva zona</p>
        <div className="flex gap-2">
          <input
            value={form.name}
            onChange={e => setForm({ ...form, name: e.target.value })}
            placeholder="Ej: Centro, La Florida…"
            className="flex-1 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-4 py-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
          />
          <input
            type="number"
            value={form.price}
            onChange={e => setForm({ ...form, price: e.target.value })}
            placeholder="Precio $"
            className="w-32 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-4 py-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
          />
          <button
            onClick={handleCreate}
            disabled={!form.name.trim() || !form.price}
            className="px-4 py-2.5 bg-[#ffd025] text-[#141414] rounded-xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            + Agregar
          </button>
        </div>
      </div>
    </div>
  );
}

function CommunityAdminPanel() {
  const { data: menu } = useGetMenu();
  const { data: menuSettings } = useGetMenu();
  const allProducts = (menu?.products ?? []) as Product[];
  const whatsapp = menuSettings?.settings?.whatsapp ?? "";
  const storeName = menuSettings?.settings?.pageTitle || "FELLAS";

  const [communitySubTab, setCommunitySubTab] = useState<"published" | "pending" | "broadcast">("published");
  const [socialSearch, setSocialSearch] = useState("");

  const normalizeStr = (s: string) =>
    s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
  const searchNorm = normalizeStr(socialSearch);

  const publishedProducts = allProducts.filter((p) => p.publishedSocial);
  const unpublishedProducts = allProducts.filter((p) => !p.publishedSocial);

  const filteredPublished = publishedProducts.filter(
    (p) => !searchNorm || normalizeStr(p.name).includes(searchNorm) || normalizeStr(p.category || "").includes(searchNorm),
  );
  const filteredUnpublished = unpublishedProducts.filter(
    (p) => !searchNorm || normalizeStr(p.name).includes(searchNorm) || normalizeStr(p.category || "").includes(searchNorm),
  );

  const [copiedId, setCopiedId] = useState<number | null>(null);
  const [copiedBroadcast, setCopiedBroadcast] = useState(false);

  const copyProductText = (p: Product) => {
    const lines = [
      `🛒 *${p.name}*`,
      `💰 $${Number(p.price).toLocaleString("es-CL")}`,
      p.category ? `📦 ${p.category}` : "",
      whatsapp ? `\n📲 Pedidos directos: wa.me/${whatsapp.replace(/\D/g, "")}` : "",
    ].filter(Boolean);
    navigator.clipboard.writeText(lines.join("\n"));
    setCopiedId(p.id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const copyBroadcastText = () => {
    const featured = allProducts.slice(0, 8);
    const text = [
      `🔥 *¡NOVEDADES Y DESTACADOS DE ${storeName.toUpperCase()}!* 🔥\n`,
      ...featured.map((p) => `• *${p.name}* — $${Number(p.price).toLocaleString("es-CL")}`),
      `\n📍 Pide directamente en nuestro menú online o escríbenos por WhatsApp:`,
      whatsapp ? `📲 wa.me/${whatsapp.replace(/\D/g, "")}` : "",
    ].filter(Boolean).join("\n");
    navigator.clipboard.writeText(text);
    setCopiedBroadcast(true);
    setTimeout(() => setCopiedBroadcast(false), 2000);
  };

  return (
    <div className="space-y-6 max-w-5xl mx-auto animate-fade-in">
      {/* Executive Header */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-xl shadow-black/60">
        <div>
          <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
            <Share2 size={15} /> Difusión & Redes Sociales
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-white uppercase">Community Manager</h2>
          <p className="text-xs text-gray-400 mt-1">
            Copia textos formateados para WhatsApp, historias de Instagram o listas de difusión.
          </p>
        </div>
        <div className="flex items-center gap-3">
          <input
            type="text"
            value={socialSearch}
            onChange={(e) => setSocialSearch(e.target.value)}
            placeholder="Buscar producto..."
            className="w-full sm:w-56 bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-xs text-white placeholder-gray-500 focus:outline-none focus:border-[#ffd025]"
          />
        </div>
      </div>

      {/* Sub-Tabs navigation */}
      <div className="flex items-center gap-2 border-b border-white/10 pb-3 overflow-x-auto scrollbar-hide">
        {[
          { id: "published", label: "Publicados en Redes", count: publishedProducts.length, icon: CheckCircle },
          { id: "pending", label: "Pendientes", count: unpublishedProducts.length, icon: Clock },
          { id: "broadcast", label: "Difusión de Catálogo", count: "WhatsApp", icon: Sparkles },
        ].map((tab) => {
          const Icon = tab.icon;
          const isActive = communitySubTab === tab.id;
          return (
            <button
              key={tab.id}
              onClick={() => setCommunitySubTab(tab.id as any)}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap border ${
                isActive
                  ? "bg-[#ffd025] text-black border-[#ffd025] shadow-lg shadow-[#ffd025]/20"
                  : "bg-white/5 text-gray-400 border-white/10 hover:text-white hover:bg-white/10"
              }`}
            >
              <Icon size={14} className={isActive ? "text-black" : "text-[#ffd025]"} />
              <span>{tab.label}</span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-mono ${isActive ? "bg-black/20 text-black font-black" : "bg-white/10 text-gray-400"}`}>
                {tab.count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Sub-Tab 1: Publicados */}
      {communitySubTab === "published" && (
        <div className="space-y-4">
          {filteredPublished.length === 0 ? (
            <div className="text-center py-12 bg-white/5 rounded-2xl border border-white/5">
              <p className="text-sm text-gray-400">No hay productos marcados como publicados en redes.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredPublished.map((p) => (
                <div key={p.id} className="bg-[#141422] rounded-2xl border border-white/10 p-3.5 flex gap-3 items-center hover:border-white/20 transition-all">
                  {p.image ? (
                    <img
                      src={p.image}
                      alt={p.name}
                      className="w-16 h-16 rounded-xl object-cover bg-black/40 shrink-0"
                      onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
                    />
                  ) : (
                    <div className="w-16 h-16 rounded-xl bg-white/5 shrink-0 flex items-center justify-center text-gray-600">
                      <ImageIcon size={20} />
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    <p className="font-bold text-sm text-white truncate">{p.name}</p>
                    <p className="text-xs text-gray-400 truncate">{p.category || "General"}</p>
                    <p className="text-sm font-black text-[#ffd025] mt-1">${Number(p.price).toLocaleString("es-CL")}</p>
                  </div>
                  <button
                    onClick={() => copyProductText(p)}
                    title="Copiar texto formateado"
                    className={`p-2.5 rounded-xl transition-all shrink-0 ${
                      copiedId === p.id
                        ? "bg-emerald-500 text-black font-bold"
                        : "bg-white/5 hover:bg-[#ffd025] hover:text-black text-gray-300 border border-white/10"
                    }`}
                  >
                    {copiedId === p.id ? <Check size={16} /> : <Copy size={16} />}
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Sub-Tab 2: Pendientes */}
      {communitySubTab === "pending" && (
        <div className="space-y-4">
          {filteredUnpublished.length === 0 ? (
            <div className="text-center py-12 bg-white/5 rounded-2xl border border-white/5">
              <p className="text-sm text-gray-400">¡Excelente! No tienes productos pendientes por difundir.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredUnpublished.map((p) => (
                <div key={p.id} className="bg-[#141422] rounded-2xl border border-white/10 p-3.5 flex gap-3 items-center hover:border-white/20 transition-all opacity-80 hover:opacity-100">
                  {p.image ? (
                    <img
                      src={p.image}
                      alt={p.name}
                      className="w-16 h-16 rounded-xl object-cover bg-black/40 shrink-0"
                      onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
                    />
                  ) : (
                    <div className="w-16 h-16 rounded-xl bg-white/5 shrink-0 flex items-center justify-center text-gray-600">
                      <ImageIcon size={20} />
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    <p className="font-bold text-sm text-white truncate">{p.name}</p>
                    <p className="text-xs text-gray-400 truncate">{p.category || "General"}</p>
                    <p className="text-sm font-black text-[#ffd025] mt-1">${Number(p.price).toLocaleString("es-CL")}</p>
                  </div>
                  <button
                    onClick={() => copyProductText(p)}
                    title="Copiar texto formateado"
                    className={`p-2.5 rounded-xl transition-all shrink-0 ${
                      copiedId === p.id
                        ? "bg-emerald-500 text-black font-bold"
                        : "bg-white/5 hover:bg-[#ffd025] hover:text-black text-gray-300 border border-white/10"
                    }`}
                  >
                    {copiedId === p.id ? <Check size={16} /> : <Copy size={16} />}
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Sub-Tab 3: Generador de Difusión */}
      {communitySubTab === "broadcast" && (
        <div className="bg-[#141422] rounded-3xl border border-white/10 p-6 space-y-5">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
                <Sparkles size={16} className="text-[#ffd025]" /> Mensaje de Difusión Automatizado
              </h3>
              <p className="text-xs text-gray-400 mt-1">
                Genera un boletín listo con tus productos top para enviar a grupos de clientes en WhatsApp.
              </p>
            </div>
            <button
              onClick={copyBroadcastText}
              className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-[#ffd025] to-[#ff9900] text-black font-black text-xs uppercase rounded-xl hover:scale-105 active:scale-95 transition-all shadow-md shadow-[#ffd025]/20"
            >
              {copiedBroadcast ? <Check size={15} /> : <Copy size={15} />}
              <span>{copiedBroadcast ? "¡Copiado!" : "Copiar Boletín"}</span>
            </button>
          </div>

          <div className="p-4 rounded-2xl bg-black/40 border border-white/10 font-mono text-xs text-gray-300 whitespace-pre-wrap leading-relaxed">
            {`🔥 *¡NOVEDADES Y DESTACADOS DE ${storeName.toUpperCase()}!* 🔥\n\n` +
              allProducts.slice(0, 8).map((p) => `• *${p.name}* — $${Number(p.price).toLocaleString("es-CL")}`).join("\n") +
              `\n\n📍 Pide directamente en nuestro menú online o escríbenos por WhatsApp:\n` +
              (whatsapp ? `📲 wa.me/${whatsapp.replace(/\D/g, "")}` : "")}
          </div>
        </div>
      )}
    </div>
  );
}

function OrdersAdminPanel({ role }: { role: "full" | "delivery" }) {
  const queryClient = useQueryClient();
  const { data: menu } = useGetMenu();
  const storeName = menu?.settings?.pageTitle || "Urban Bite";
  const allProducts = (menu?.products ?? []) as Product[];
  const { data: orders = [], isLoading, refetch: refetchOrders } = useListOrders();

  const patchOrderCache = (id: number, patch: Partial<Order>) =>
    queryClient.setQueryData(getListOrdersQueryKey(), (old: any) =>
      Array.isArray(old) ? old.map((o: Order) => (o.id === id ? { ...o, ...patch } : o)) : old,
    );

  useEffect(() => {
    const interval = setInterval(() => {
      refetchOrders();
    }, 15_000);
    return () => clearInterval(interval);
  }, [refetchOrders]);
  const deleteOrderMut = useDeleteOrder();
  const updateStatusMut = useUpdateOrderStatus();

  type EditItem = { name: string; quantity: number; price: number; selectedOption?: string };
  const [statusFilter, setStatusFilter] = useState<string>("Todos");
  const [orderSearch, setOrderSearch] = useState<string>("");
  const [expandedOrders, setExpandedOrders] = useState<Record<number, boolean>>({});
  const [editingOrder, setEditingOrder] = useState<Order | null>(null);
  const [editItems, setEditItems] = useState<EditItem[]>([]);
  const [editSearch, setEditSearch] = useState("");
  const [editNote, setEditNote] = useState("");
  const [panelToast, setPanelToast] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setPanelToast(msg);
    setTimeout(() => setPanelToast(null), 3000);
  };

  const getNextStatus = (currentStatus: string, address?: string) => {
    const isRetiro = address?.toLowerCase().includes("retiro") ?? false;
    if (currentStatus === "Pendiente" || !currentStatus) return "Confirmado y en Preparación";
    if (currentStatus === "Confirmado y en Preparación") return isRetiro ? "Listo para Retirar" : "Delivery en Camino";
    if (currentStatus === "Delivery en Camino" || currentStatus === "Listo para Retirar") return "Entregado";
    return null;
  };

  const openEdit = (order: Order) => {
    const items = (order.items ?? []) as EditItem[];
    setEditItems(items.map((it) => ({ ...it })));
    setEditNote((order as Order & { modificationNote?: string | null }).modificationNote ?? "");
    setEditSearch("");
    setEditingOrder(order);
  };

  const handleSaveEdit = () => {
    if (!editingOrder) return;
    const newTotal = editItems.reduce((s, it) => s + it.price * it.quantity, 0);
    patchOrderCache(editingOrder.id, { items: editItems as any, total: String(newTotal), modificationNote: editNote || null } as any);
    setEditingOrder(null);
    updateStatusMut.mutate(
      {
        id: editingOrder.id,
        data: {
          status: editingOrder.status,
          items: editItems,
          modificationNote: editNote || undefined,
          total: newTotal,
        } as Parameters<typeof updateStatusMut.mutate>[0]["data"],
      },
      {
        onError: () => {
          refetchOrders();
          showToast("Error al guardar los cambios. Intenta de nuevo.");
        },
      },
    );
  };

  const handleSetCaja = (order: Order, caja: string | null) => {
    patchOrderCache(order.id, { caja: caja ?? undefined } as any);
    updateStatusMut.mutate({
      id: order.id,
      data: { status: order.status, caja: caja ?? undefined } as Parameters<typeof updateStatusMut.mutate>[0]["data"],
    }, { onError: () => refetchOrders() });
  };

  const knownIdsRef = useRef<Set<number> | null>(null);
  const notifPermissionRef = useRef<NotificationPermission>("default");

  useEffect(() => {
    if ("Notification" in window && Notification.permission === "default") {
      Notification.requestPermission().then((perm) => {
        notifPermissionRef.current = perm;
      });
    } else if ("Notification" in window) {
      notifPermissionRef.current = Notification.permission;
    }
  }, []);

  useEffect(() => {
    if (isLoading || orders.length === 0) return;
    const currentIds = new Set(orders.map((o: Order) => o.id));

    if (knownIdsRef.current === null) {
      knownIdsRef.current = currentIds;
      return;
    }

    const newOrders = orders.filter(
      (o: Order) => !knownIdsRef.current!.has(o.id),
    );
    knownIdsRef.current = currentIds;

    if (newOrders.length === 0) return;

    if (
      "Notification" in window &&
      (notifPermissionRef.current === "granted" ||
        Notification.permission === "granted")
    ) {
      newOrders.forEach((o: Order) => {
        new Notification(`🛒 Nuevo pedido — ${storeName}`, {
          body: `${o.customerName} · $${Number(o.total).toLocaleString("es-CL")}`,
          icon: "/favicon.ico",
        });
      });
    }
  }, [orders, isLoading, storeName]);

  const handleDelete = (id: number) => {
    if (!confirm("¿Eliminar este pedido del registro?")) return;
    queryClient.setQueryData(getListOrdersQueryKey(), (old: any) =>
      Array.isArray(old) ? old.filter((o: Order) => o.id !== id) : old,
    );
    deleteOrderMut.mutate({ id }, { onError: () => refetchOrders() });
  };

  const handleAnular = (id: number) => {
    try {
      if (typeof window !== "undefined" && window.confirm && !window.confirm("¿Anular este pedido? Quedará registrado pero no contará en las ventas.")) return;
    } catch {}
    patchOrderCache(id, { status: "Anulado" });
    updateStatusMut.mutate(
      { id, data: { status: "Anulado" } },
      { onError: () => { refetchOrders(); showToast("Error al anular el pedido. Intenta de nuevo."); } },
    );
  };

  const buildStatusMessage = (customerName: string, status: string) => {
    const first = customerName.split(" ")[0] || customerName;
    switch (status) {
      case "Confirmado y en Preparación":
        return `¡Hola ${first}! 👋 Tu pedido en *${storeName}* fue *confirmado* y ya lo estamos alistando 🍾\n\nEn breve te avisamos cuando esté listo.`;
      case "Listo para Retirar":
        return `¡Hola ${first}! 🍻 Tu pedido en *${storeName}* ya está *listo para retirar* 🏪✅\n\nTe esperamos en el local con todo listo. ¡Salud! 🥃`;
      case "Delivery en Camino":
        return `¡Hola ${first}! 🛵 Tu pedido de *${storeName}* va *en camino* con nuestro delivery 🍾\n\nEn unos minutos llega a tu dirección. ¡Gracias por tu preferencia! 🥂🎊`;
      default:
        return "";
    }
  };

  const handleSetStatus = (
    id: number,
    status: string,
    customerName: string,
    phone: string,
    sendsWhatsapp: boolean,
  ) => {
    if (sendsWhatsapp) {
      const cleanedPhone = (phone ?? "").replace(/\D/g, "");
      if (!cleanedPhone) {
        showToast("Este pedido no tiene número de teléfono. No se puede enviar el mensaje.");
        return;
      }
      const message = buildStatusMessage(customerName, status);
      try {
        window.open(
          `https://api.whatsapp.com/send?phone=${cleanedPhone}&text=${encodeURIComponent(message)}`,
          "_blank",
        );
      } catch {}
    }
    patchOrderCache(id, { status });
    updateStatusMut.mutate(
      { id, data: { status } },
      { onError: () => { refetchOrders(); showToast("Error al actualizar el estado. Intenta de nuevo."); } },
    );
  };

  const normalizeStr = (s: string) =>
    s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");

  return (
    <>
    {editingOrder && (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
        <div className="bg-[#1a1a1a] rounded-3xl border border-[#ffd025]/30 w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl">
          <div className="flex items-center justify-between p-5 border-b border-gray-800">
            <h3 className="text-base font-black uppercase flex items-center gap-2">
              <Pencil size={16} className="text-[#ffd025]" />
              Editar Pedido #{editingOrder.id} — {editingOrder.customerName}
            </h3>
            <button
              type="button"
              onClick={() => setEditingOrder(null)}
              className="p-2 text-gray-500 hover:text-white hover:bg-gray-800 rounded-xl transition-colors"
            >
              <X size={16} />
            </button>
          </div>
          <div className="flex-1 overflow-y-auto p-5 space-y-5">
            <div>
              <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                Productos del pedido
              </p>
              <ul className="space-y-1.5">
                {editItems.map((it, idx) => (
                  <li
                    key={idx}
                    className="flex items-center gap-3 bg-[#141414] border border-gray-800 rounded-xl px-3 py-2"
                  >
                    <div className="flex items-center gap-1.5 mr-1">
                      <button
                        type="button"
                        onClick={() =>
                          setEditItems(
                            editItems.map((x, i) =>
                              i === idx ? { ...x, quantity: Math.max(1, x.quantity - 1) } : x,
                            ),
                          )
                        }
                        className="w-5 h-5 rounded-full bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold flex items-center justify-center transition-colors"
                      >
                        −
                      </button>
                      <span className="text-[#ffd025] font-black text-sm w-5 text-center">
                        {it.quantity}
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          setEditItems(
                            editItems.map((x, i) =>
                              i === idx ? { ...x, quantity: x.quantity + 1 } : x,
                            ),
                          )
                        }
                        className="w-5 h-5 rounded-full bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold flex items-center justify-center transition-colors"
                      >
                        +
                      </button>
                    </div>
                    <span className="flex-1 text-sm text-white">
                      {it.name}
                      {it.selectedOption && (
                        <span className="text-gray-500"> ({it.selectedOption})</span>
                      )}
                    </span>
                    <span className="text-gray-400 text-sm">
                      ${(it.price * it.quantity).toLocaleString("es-CL")}
                    </span>
                    <button
                      type="button"
                      onClick={() => setEditItems(editItems.filter((_, i) => i !== idx))}
                      className="p-1 text-red-500 hover:text-red-400 hover:bg-red-500/10 rounded-lg transition-colors"
                      title="Quitar producto"
                    >
                      <X size={13} />
                    </button>
                  </li>
                ))}
              </ul>
              {editItems.length === 0 && (
                <p className="text-gray-600 text-sm text-center py-4">
                  Sin productos — agrega uno desde el buscador
                </p>
              )}
              {editItems.length > 0 && (
                <div className="mt-2 text-right text-sm">
                  <span className="text-gray-500">Nuevo total: </span>
                  <span className="text-[#ffd025] font-black">
                    ${editItems.reduce((s, it) => s + it.price * it.quantity, 0).toLocaleString("es-CL")}
                  </span>
                </div>
              )}
            </div>

            <div>
              <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                Agregar / reemplazar producto
              </p>
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none">
                  🔍
                </span>
                <input
                  type="text"
                  placeholder="Buscar en el catálogo..."
                  value={editSearch}
                  onChange={(e) => setEditSearch(e.target.value)}
                  className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl pl-9 py-2 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder:text-gray-600"
                />
              </div>
              {editSearch.trim() && (() => {
                const q = normalizeStr(editSearch.trim());
                const words = q.split(/\s+/).filter(Boolean);
                const results = allProducts
                  .filter((p) => {
                    const haystack = normalizeStr(
                      [p.name, p.category, p.aisle, p.subcategory ?? ""].join(" "),
                    );
                    return words.every((w) => haystack.includes(w));
                  })
                  .slice(0, 12);
                return (
                  <div className="mt-1.5 max-h-44 overflow-y-auto space-y-1 rounded-xl border border-gray-800 bg-[#141414] p-1">
                    {results.length === 0 ? (
                      <p className="text-gray-600 text-sm text-center py-3">Sin resultados</p>
                    ) : (
                      results.map((p) => (
                        <button
                          key={p.id}
                          type="button"
                          onClick={() => {
                            setEditItems([
                              ...editItems,
                              { name: p.name, quantity: 1, price: p.price },
                            ]);
                            setEditSearch("");
                          }}
                          className="w-full text-left flex items-center justify-between gap-2 hover:bg-[#222] rounded-lg px-3 py-1.5 transition-colors"
                        >
                          <span className="text-sm text-white">{p.name}</span>
                          <span className="text-[#ffd025] font-bold text-sm whitespace-nowrap">
                            ${p.price.toLocaleString("es-CL")}
                          </span>
                        </button>
                      ))
                    )}
                  </div>
                );
              })()}
            </div>

            <div>
              <label className="text-[10px] uppercase font-bold text-gray-500 mb-2 block tracking-widest">
                Motivo del cambio (opcional)
              </label>
              <textarea
                rows={2}
                placeholder="Ej: No había Cerveza Austral, se reemplazó por Cristal 500ml..."
                value={editNote}
                onChange={(e) => setEditNote(e.target.value)}
                className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025] focus:outline-none resize-none placeholder:text-gray-600"
              />
            </div>
          </div>
          <div className="flex gap-3 p-5 border-t border-gray-800">
            <button
              type="button"
              onClick={() => setEditingOrder(null)}
              className="flex-1 py-2.5 rounded-xl border border-gray-700 text-gray-400 hover:text-white hover:border-gray-500 text-sm font-bold transition-colors"
            >
              Cancelar
            </button>
            <button
              type="button"
              onClick={handleSaveEdit}
              disabled={updateStatusMut.isPending || editItems.length === 0}
              className="flex-1 py-2.5 bg-[#ffd025] text-[#141414] rounded-xl font-bold text-sm uppercase hover:bg-[#e5b81a] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Guardar cambios
            </button>
          </div>
        </div>
      </div>
    )}
    <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-8 rounded-3xl border border-white/10 animate-fade-in space-y-6 shadow-2xl shadow-black/80">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-white/10">
        <div>
          <h2 className="text-xl sm:text-2xl font-black uppercase text-white flex items-center gap-2">
            <ClipboardList size={22} className="text-[#ffd025]" /> Comandas y Pedidos
            <span className="ml-2 text-xs font-bold text-[#ffd025] bg-[#ffd025]/10 px-2.5 py-1 rounded-full border border-[#ffd025]/20">
              {orders.length} totales
            </span>
          </h2>
          <p className="text-xs text-gray-400 mt-1">Gestión de comandas en vivo con avance paso a paso y tarjetas plegables.</p>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 text-xs text-gray-400 bg-white/5 px-3 py-1.5 rounded-full border border-white/5">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
            <span>En vivo</span>
          </div>
          <button
            type="button"
            onClick={() => {
              const anyOpen = Object.values(expandedOrders).some(Boolean);
              const nextState: Record<number, boolean> = {};
              orders.forEach((o: Order) => {
                nextState[o.id] = !anyOpen;
              });
              setExpandedOrders(nextState);
            }}
            className="px-3.5 py-2 bg-white/5 hover:bg-white/10 text-gray-300 rounded-xl text-xs font-bold border border-white/10 flex items-center gap-2 transition-colors"
          >
            <Layers size={13} />
            {Object.values(expandedOrders).some(Boolean) ? "Plegar Comandas" : "Desplegar Todas"}
          </button>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
        {/* Filter Status Tabs */}
        <div className="flex items-center gap-2 overflow-x-auto pb-1 border-b md:border-b-0 border-white/5 scrollbar-hide">
          {[
            { label: "Todos", value: "Todos" },
            { label: "Pendientes", value: "Pendiente" },
            { label: "En Preparación", value: "Confirmado y en Preparación" },
            { label: "En Camino", value: "Delivery en Camino" },
            { label: "Listos para Retirar", value: "Listo para Retirar" },
            { label: "Entregados", value: "Entregado" },
            { label: "Anulados", value: "Anulado" },
          ].map((tab) => {
            const count =
              tab.value === "Todos"
                ? orders.length
                : orders.filter((o: Order) => (tab.value === "Pendiente" ? (o.status === "Pendiente" || !o.status) : o.status === tab.value)).length;
            const isActive = statusFilter === tab.value;
            return (
              <button
                key={tab.value}
                type="button"
                onClick={() => setStatusFilter(tab.value)}
                className={`flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-bold uppercase transition-all whitespace-nowrap shrink-0 border ${
                  isActive
                    ? "bg-[#ffd025] text-[#141414] border-[#ffd025] shadow-md shadow-[#ffd025]/20"
                    : "bg-[#181826] text-gray-400 hover:text-white border-white/5"
                }`}
              >
                <span>{tab.label}</span>
                <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-black ${isActive ? "bg-[#141414]/20 text-[#141414]" : "bg-white/10 text-gray-300"}`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* Search Orders */}
        <div className="relative flex-1 max-w-sm">
          <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none">
            🔍
          </span>
          <input
            type="text"
            placeholder="Buscar pedido por cliente, teléfono, dirección..."
            value={orderSearch}
            onChange={(e) => setOrderSearch(e.target.value)}
            className="w-full bg-[#181826] border border-white/10 rounded-xl pl-9 pr-9 py-2 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-500"
          />
          {orderSearch && (
            <button
              type="button"
              onClick={() => setOrderSearch("")}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-white"
            >
              ✕
            </button>
          )}
        </div>
      </div>

      {isLoading ? (
        <p className="text-gray-500 text-sm">Cargando pedidos...</p>
      ) : orders.length === 0 ? (
        <div className="text-center py-12 opacity-60">
          <ClipboardList size={40} className="mx-auto mb-3 text-[#ffd025]/40" />
          <p className="text-sm uppercase font-bold text-gray-400">
            Aún no hay pedidos registrados
          </p>
        </div>
      ) : (
        <div className="space-y-4">
          {orders
            .filter((o: Order) => {
              if (statusFilter === "Todos") return true;
              if (statusFilter === "Pendiente") return o.status === "Pendiente" || !o.status;
              return o.status === statusFilter;
            })
            .filter((o: Order) => {
              if (!orderSearch.trim()) return true;
              const q = orderSearch.toLowerCase();
              return (
                o.customerName?.toLowerCase().includes(q) ||
                o.phone?.toLowerCase().includes(q) ||
                o.address?.toLowerCase().includes(q) ||
                String(o.id).includes(q)
              );
            })
            .map((order: Order) => {
            const date = new Date(order.createdAt);
            const items = (order.items ?? []) as Array<{
              name: string;
              quantity: number;
              price: number;
              selectedOption?: string;
            }>;
            const isExpanded = expandedOrders[order.id] ?? (order.status === "Pendiente" || !order.status);
            const nextStatus = getNextStatus(order.status, order.address);

            return (
              <div
                key={order.id}
                className={`border rounded-2xl p-5 transition-all shadow-lg ${
                  order.status === "Anulado"
                    ? "bg-[#181116]/80 border-red-900/40 opacity-70"
                    : order.status === "Entregado"
                    ? "bg-[#101a14]/80 border-green-700/30 opacity-80"
                    : "bg-[#181828] border-white/10 hover:border-[#ffd025]/40"
                }`}
              >
                {/* Header Summary Row */}
                <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-4">
                  <div className="space-y-1">
                    <div className="flex items-center gap-2 text-white font-black text-base flex-wrap">
                      <User size={16} className="text-[#ffd025]" />
                      <span>{order.customerName}</span>
                      <span className="text-xs font-bold text-gray-500">#{order.id}</span>
                      {order.status === "Anulado" && (
                        <span className="text-[10px] font-black uppercase text-red-400 bg-red-600/15 border border-red-600/30 px-2 py-0.5 rounded-full flex items-center gap-1">
                          <Ban size={10} /> Anulado
                        </span>
                      )}
                      {order.status === "Entregado" && (
                        <span className="text-[10px] font-black uppercase text-green-400 bg-green-600/15 border border-green-600/30 px-2 py-0.5 rounded-full">
                          ✓ Entregado
                        </span>
                      )}
                      {(!order.status || order.status === "Pendiente") && (
                        <span className="text-[10px] font-black uppercase text-amber-400 bg-amber-500/15 border border-amber-500/30 px-2 py-0.5 rounded-full">
                          ⏳ Pendiente
                        </span>
                      )}
                      {order.status === "Confirmado y en Preparación" && (
                        <span className="text-[10px] font-black uppercase text-blue-400 bg-blue-500/15 border border-blue-500/30 px-2 py-0.5 rounded-full">
                          🍳 En Preparación
                        </span>
                      )}
                      {order.status === "Delivery en Camino" && (
                        <span className="text-[10px] font-black uppercase text-cyan-400 bg-cyan-500/15 border border-cyan-500/30 px-2 py-0.5 rounded-full">
                          🛵 En Camino
                        </span>
                      )}
                      {order.status === "Listo para Retirar" && (
                        <span className="text-[10px] font-black uppercase text-purple-400 bg-purple-500/15 border border-purple-500/30 px-2 py-0.5 rounded-full">
                          🏪 Listo para Retirar
                        </span>
                      )}
                    </div>

                    <div className="flex items-start gap-2 text-xs text-gray-300">
                      <MapPin size={13} className="text-[#ffd025] mt-0.5 shrink-0" />
                      <span>{order.address}</span>
                    </div>

                    {order.phone && (
                      <div className="flex items-center gap-2 text-xs text-gray-400">
                        <Phone size={12} className="text-[#ffd025]" />
                        <span>{order.phone}</span>
                      </div>
                    )}

                    <p className="text-[10px] uppercase tracking-wider text-gray-500">
                      {date.toLocaleString("es-CL", {
                        day: "2-digit",
                        month: "2-digit",
                        year: "numeric",
                        hour: "2-digit",
                        minute: "2-digit",
                      })}
                    </p>
                  </div>

                  {/* Actions & Total */}
                  <div className="flex flex-col sm:items-end gap-2.5">
                    <div className="flex items-baseline gap-2">
                      <span className="text-xl sm:text-2xl font-black text-[#ffd025]">
                        ${order.total.toLocaleString("es-CL")}
                      </span>
                    </div>

                    {order.discountCode && order.discountAmount && order.subtotal && (
                      <div className="flex items-center gap-1.5 text-[11px]">
                        <span className="text-green-400 font-bold bg-green-500/10 border border-green-500/25 px-2 py-0.5 rounded-full">
                          🏷️ {order.discountCode} (-${order.discountAmount.toLocaleString("es-CL")})
                        </span>
                      </div>
                    )}

                    {/* Step Advance & Collapse Controls */}
                    <div className="flex items-center gap-2 flex-wrap">
                      {nextStatus && order.status !== "Anulado" && (
                        <button
                          type="button"
                          onClick={() => handleSetStatus(order.id, nextStatus, order.customerName, order.phone ?? "", false)}
                          disabled={updateStatusMut.isPending}
                          className="px-3.5 py-1.5 bg-gradient-to-r from-[#ffd025] to-[#e6b800] text-[#141414] font-black text-xs rounded-xl shadow-md hover:scale-[1.02] transition-all flex items-center gap-1 disabled:opacity-50"
                        >
                          <span>Avanzar ➔</span>
                        </button>
                      )}

                      <button
                        type="button"
                        onClick={() => setExpandedOrders((prev) => ({ ...prev, [order.id]: !isExpanded }))}
                        className="px-3 py-1.5 rounded-xl bg-white/5 hover:bg-white/10 text-gray-300 font-bold text-xs border border-white/10 flex items-center gap-1.5 transition-colors"
                      >
                        {isExpanded ? (
                          <>
                            <ChevronUp size={13} /> Plegar Detalle
                          </>
                        ) : (
                          <>
                            <ChevronDown size={13} /> Ver ({items.length} prod.)
                          </>
                        )}
                      </button>

                      {order.status !== "Anulado" && (
                        <button
                          onClick={() => openEdit(order)}
                          className="p-1.5 text-gray-400 hover:text-[#ffd025] hover:bg-[#ffd025]/10 rounded-lg transition-colors"
                          aria-label="Editar pedido"
                          title="Editar productos del pedido"
                        >
                          <Pencil size={15} />
                        </button>
                      )}

                      {order.status !== "Anulado" && (
                        <button
                          onClick={() => handleAnular(order.id)}
                          disabled={updateStatusMut.isPending}
                          className="p-1.5 text-gray-400 hover:text-orange-400 hover:bg-orange-500/10 rounded-lg transition-colors disabled:opacity-40"
                          aria-label="Anular pedido"
                          title="Anular pedido"
                        >
                          <Ban size={15} />
                        </button>
                      )}

                      {role === "full" && (
                        <button
                          onClick={() => handleDelete(order.id)}
                          className="p-1.5 text-gray-400 hover:text-red-500 hover:bg-red-500/10 rounded-lg transition-colors"
                          aria-label="Eliminar pedido"
                        >
                          <Trash2 size={15} />
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Collapsible Details Body */}
                {isExpanded && (
                  <div className="mt-4 pt-4 border-t border-white/10 space-y-4 animate-fade-in">
                    {order.status !== "Anulado" && (
                      <div className="flex items-center gap-2 flex-wrap text-xs">
                        <CreditCard size={13} className="text-gray-400" />
                        <span className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Cobrado desde:</span>
                        {[null, "Caja 1", "Caja 2"].map((c) => {
                          const orderCaja = (order as Order & { caja?: string | null }).caja ?? null;
                          const isActive = orderCaja === c;
                          return (
                            <button
                              key={c ?? "ninguna"}
                              type="button"
                              onClick={() => handleSetCaja(order, c)}
                              disabled={updateStatusMut.isPending}
                              className={`text-[10px] px-2.5 py-0.5 rounded-full font-bold border transition-colors disabled:opacity-50 ${
                                isActive
                                  ? c === null
                                    ? "bg-gray-700 border-gray-500 text-gray-300"
                                    : c === "Caja 1"
                                    ? "bg-blue-600 border-blue-400 text-white"
                                    : "bg-purple-600 border-purple-400 text-white"
                                  : "bg-transparent border-white/10 text-gray-400 hover:border-white/30"
                              }`}
                            >
                              {c ?? "Sin asignar"}
                            </button>
                          );
                        })}
                      </div>
                    )}

                    <div className="bg-[#12121e] p-3.5 rounded-xl border border-white/5">
                      <p className="text-[10px] uppercase font-bold text-gray-400 mb-2 tracking-wider">
                        Productos Solicitados ({items.reduce((s, it) => s + it.quantity, 0)} unidades)
                      </p>
                      <ul className="space-y-1.5 divide-y divide-white/5">
                        {items.map((it, idx) => (
                          <li key={idx} className="flex justify-between items-center text-xs sm:text-sm text-gray-300 pt-1.5 first:pt-0">
                            <span>
                              <span className="text-[#ffd025] font-black mr-2 bg-[#ffd025]/10 px-1.5 py-0.5 rounded text-xs">
                                {it.quantity}x
                              </span>
                              <span className="text-white font-medium">{it.name}</span>
                              {it.selectedOption && (
                                <span className="text-gray-400 text-xs ml-1.5">({it.selectedOption})</span>
                              )}
                            </span>
                            <span className="text-gray-300 font-bold whitespace-nowrap">
                              ${(it.price * it.quantity).toLocaleString("es-CL")}
                            </span>
                          </li>
                        ))}
                      </ul>
                    </div>

                    {order.notes && (
                      <div className="p-3 bg-white/[0.02] border border-white/5 rounded-xl text-xs text-gray-300">
                        <span className="font-bold text-gray-400 uppercase mr-2 text-[10px]">Notas del cliente:</span>
                        {order.notes}
                      </div>
                    )}

                    {(order as Order & { modificationNote?: string | null }).modificationNote && (
                      <div className="p-3 border border-amber-900/40 text-xs text-amber-300 bg-amber-950/20 rounded-xl">
                        <span className="font-bold text-amber-400 uppercase mr-2 text-[10px]">✏️ Cambio realizado:</span>
                        {(order as Order & { modificationNote?: string | null }).modificationNote}
                      </div>
                    )}

                    {order.status !== "Anulado" && (
                      <div className="pt-2">
                        <p className="text-[10px] uppercase font-bold text-gray-400 mb-2 tracking-wider">
                          Cambiar Estado & Notificar por WhatsApp
                        </p>
                        <div className="flex flex-wrap gap-2">
                          {ORDER_STATUSES.map((s) => {
                            const isActive = order.status === s.value;
                            return (
                              <button
                                key={s.value}
                                type="button"
                                onClick={() =>
                                  handleSetStatus(
                                    order.id,
                                    s.value,
                                    order.customerName,
                                    order.phone ?? "",
                                    s.whatsapp,
                                  )
                                }
                                disabled={updateStatusMut.isPending}
                                className={`text-[11px] font-bold uppercase tracking-wide px-3 py-1.5 rounded-xl border transition-all disabled:opacity-50 ${
                                  isActive
                                    ? s.activeColor + " shadow-md"
                                    : s.color + " hover:brightness-125"
                                }`}
                              >
                                {s.label}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
    </>
  );
}

const DOW_NAMES = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

function StatsAdminPanel() {
  const [period, setPeriod] = useState<"weekly" | "monthly">("weekly");
  const [statsSubTab, setStatsSubTab] = useState<"sales" | "traffic" | "top">("sales");
  const [isGenerating, setIsGenerating] = useState(false);
  const visitsQ = useGetVisitStats();
  const salesQ = useGetSalesStats({ period });
  const { data: menu } = useGetMenu();
  const storeName = menu?.settings?.pageTitle || "FELLAS";

  const visits = visitsQ.data;
  const sales = salesQ.data;

  const handleDownloadPDF = async () => {
    setIsGenerating(true);
    try {
      const [freshVisits, freshSales] = await Promise.all([
        visitsQ.refetch(),
        salesQ.refetch(),
      ]);
      generateStatsPDF(storeName, period, freshVisits.data, freshSales.data);
    } finally {
      setIsGenerating(false);
    }
  };

  const maxVisitHour = visits
    ? Math.max(1, ...visits.hourlyDistribution.map((b) => b.count))
    : 1;
  const maxSalesHour = sales
    ? Math.max(1, ...sales.hourlyDistribution.map((b) => b.orders))
    : 1;
  const maxDow = sales
    ? Math.max(1, ...sales.dayOfWeekDistribution.map((b) => b.orders))
    : 1;
  const maxTopQty = sales?.topProducts.length
    ? Math.max(1, ...sales.topProducts.map((p) => p.qty))
    : 1;

  const peakVisitHour = visits?.hourlyDistribution.reduce(
    (best, cur) => (cur.count > best.count ? cur : best),
    { hour: 0, count: 0 },
  );
  const peakSalesHour = sales?.hourlyDistribution.reduce(
    (best, cur) => (cur.orders > best.orders ? cur : best),
    { hour: 0, orders: 0, revenue: 0 },
  );
  const peakDow = sales?.dayOfWeekDistribution.reduce(
    (best, cur) => (cur.orders > best.orders ? cur : best),
    { dayOfWeek: 0, orders: 0, revenue: 0 },
  );

  return (
    <div className="space-y-6 max-w-5xl mx-auto animate-fade-in">
      {/* Header with PDF Download & Period Switcher */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-xl shadow-black/60">
        <div>
          <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
            <BarChart3 size={15} /> Estadísticas & Rendimiento
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-white uppercase">Métricas del Negocio</h2>
          <p className="text-xs text-gray-400 mt-0.5">
            Analiza el comportamiento de compra, horarios con mayor demanda y tráfico web.
          </p>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={handleDownloadPDF}
            disabled={isGenerating || (!visits && !sales)}
            className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-[#ffd025] via-[#ffda47] to-[#e6b800] text-[#0a0a0f] rounded-xl font-black uppercase text-xs hover:scale-[1.01] active:scale-[0.99] transition-all shadow-lg shadow-[#ffd025]/20 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <Download size={15} />
            <span>{isGenerating ? "Generando..." : "Descargar Infografía PDF"}</span>
          </button>
        </div>
      </div>

      {/* Sub-Tabs navigation */}
      <div className="flex items-center gap-2 border-b border-white/10 pb-3 overflow-x-auto scrollbar-hide">
        {[
          { id: "sales", label: "1. Ventas & Facturación", count: sales ? `$${sales.totalRevenue.toLocaleString("es-CL")}` : "0", icon: TrendingUp },
          { id: "traffic", label: "2. Visitas & Tráfico Web", count: visits ? `${visits.total} visitas` : "0", icon: Eye },
          { id: "top", label: "3. Lo Más Vendido", count: sales ? `${sales.topProducts.length} ítems` : "0", icon: Flame },
        ].map((tab) => {
          const Icon = tab.icon;
          const isActive = statsSubTab === tab.id;
          return (
            <button
              key={tab.id}
              onClick={() => setStatsSubTab(tab.id as any)}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap border ${
                isActive
                  ? "bg-[#ffd025] text-black border-[#ffd025] shadow-lg shadow-[#ffd025]/20"
                  : "bg-white/5 text-gray-400 border-white/10 hover:text-white hover:bg-white/10"
              }`}
            >
              <Icon size={14} className={isActive ? "text-black" : "text-[#ffd025]"} />
              <span>{tab.label}</span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-mono ${isActive ? "bg-black/20 text-black font-black" : "bg-white/10 text-gray-400"}`}>
                {tab.count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Tab 1: Ventas & Facturación */}
      {statsSubTab === "sales" && (
        <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-8 rounded-3xl border border-white/10 shadow-2xl shadow-black/80 space-y-6">
          <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-white/10">
            <div>
              <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
                <TrendingUp size={18} className="text-[#ffd025]" /> Rendimiento de Pedidos
              </h3>
              <p className="text-xs text-gray-400 mt-0.5">
                {period === "weekly" ? "Últimos 7 días de operación" : "Últimos 30 días de operación"}
              </p>
            </div>
            <div className="inline-flex bg-black/40 rounded-xl p-1 border border-white/10">
              <button
                onClick={() => setPeriod("weekly")}
                className={`px-3.5 py-1.5 text-xs font-bold uppercase rounded-lg transition-all ${
                  period === "weekly"
                    ? "bg-[#ffd025] text-black shadow-sm font-black"
                    : "text-gray-400 hover:text-white"
                }`}
              >
                Semanal
              </button>
              <button
                onClick={() => setPeriod("monthly")}
                className={`px-3.5 py-1.5 text-xs font-bold uppercase rounded-lg transition-all ${
                  period === "monthly"
                    ? "bg-[#ffd025] text-black shadow-sm font-black"
                    : "text-gray-400 hover:text-white"
                }`}
              >
                Mensual
              </button>
            </div>
          </div>

          {salesQ.isLoading || !sales ? (
            <p className="text-gray-500 text-sm py-10 text-center">Cargando reporte de ventas...</p>
          ) : (
            <>
              <div className="grid grid-cols-2 gap-3.5">
                <StatCard
                  label="Pedidos entregados"
                  value={sales.totalOrders}
                  icon={<ClipboardList size={16} />}
                />
                <StatCard
                  label="Ingresos Totales"
                  value={`$${sales.totalRevenue.toLocaleString("es-CL")}`}
                  icon={<TrendingUp size={16} />}
                />
              </div>

              {sales.cancelledOrders > 0 && (
                <div className="flex items-center gap-2 px-3.5 py-2.5 rounded-xl border border-red-500/20 bg-red-500/10 text-xs text-red-300">
                  <Ban size={14} className="shrink-0" />
                  <span>
                    <strong className="text-red-200">{sales.cancelledOrders}</strong>{" "}
                    {sales.cancelledOrders === 1 ? "pedido anulado" : "pedidos anulados"} en este período (no sumados a ingresos).
                  </span>
                </div>
              )}

              <div className="grid sm:grid-cols-2 gap-3 text-xs uppercase tracking-wider text-gray-400">
                {peakSalesHour && peakSalesHour.orders > 0 && (
                  <div className="bg-[#161626] p-3.5 rounded-2xl border border-white/5 flex items-center justify-between">
                    <span>Hora con más pedidos:</span>
                    <span className="text-[#ffd025] font-black text-sm">
                      {String(peakSalesHour.hour).padStart(2, "0")}:00 · {peakSalesHour.orders} pedidos
                    </span>
                  </div>
                )}
                {peakDow && peakDow.orders > 0 && (
                  <div className="bg-[#161626] p-3.5 rounded-2xl border border-white/5 flex items-center justify-between">
                    <span>Día con más demanda:</span>
                    <span className="text-[#ffd025] font-black text-sm">
                      {DOW_NAMES[peakDow.dayOfWeek]} · {peakDow.orders} pedidos
                    </span>
                  </div>
                )}
              </div>

              <div className="grid md:grid-cols-2 gap-6 pt-2">
                <div className="bg-[#161626] p-4 rounded-2xl border border-white/5 space-y-3">
                  <h4 className="text-xs uppercase font-bold text-gray-300 flex items-center gap-2">
                    <Clock size={14} className="text-[#ffd025]" /> Pedidos por hora
                  </h4>
                  <div className="flex items-end gap-1 h-32 bg-black/30 rounded-xl p-3 border border-white/5">
                    {sales.hourlyDistribution.map((b) => (
                      <div key={b.hour} className="flex-1 flex flex-col items-center gap-1 group">
                        <div className="text-[9px] text-gray-500 group-hover:text-[#ffd025]">
                          {b.orders > 0 ? b.orders : ""}
                        </div>
                        <div
                          className="w-full bg-gradient-to-t from-orange-500/40 to-[#ffd025] rounded-t transition-all"
                          style={{
                            height: `${(b.orders / maxSalesHour) * 85}%`,
                            minHeight: b.orders > 0 ? 4 : 0,
                          }}
                          title={`${String(b.hour).padStart(2, "0")}:00 - ${b.orders} pedidos`}
                        />
                        <div className="text-[9px] text-gray-500">{b.hour}</div>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="bg-[#161626] p-4 rounded-2xl border border-white/5 space-y-3">
                  <h4 className="text-xs uppercase font-bold text-gray-300 flex items-center gap-2">
                    <Calendar size={14} className="text-[#ffd025]" /> Pedidos por día de la semana
                  </h4>
                  <div className="flex items-end gap-2 h-32 bg-black/30 rounded-xl p-3 border border-white/5">
                    {sales.dayOfWeekDistribution.map((b) => (
                      <div
                        key={b.dayOfWeek}
                        className="flex-1 flex flex-col items-center gap-1 group"
                      >
                        <div className="text-[9px] text-gray-500 group-hover:text-[#ffd025]">
                          {b.orders > 0 ? b.orders : ""}
                        </div>
                        <div
                          className="w-full bg-gradient-to-t from-orange-500/40 to-[#ffd025] rounded-t transition-all"
                          style={{
                            height: `${(b.orders / maxDow) * 85}%`,
                            minHeight: b.orders > 0 ? 4 : 0,
                          }}
                          title={`${DOW_NAMES[b.dayOfWeek]} - ${b.orders} pedidos`}
                        />
                        <div className="text-[10px] text-gray-400 font-bold">
                          {DOW_NAMES[b.dayOfWeek].slice(0, 3)}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* Tab 2: Visitas & Tráfico Web */}
      {statsSubTab === "traffic" && (
        <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-8 rounded-3xl border border-white/10 shadow-2xl shadow-black/80 space-y-6">
          <div>
            <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
              <Eye size={18} className="text-[#ffd025]" /> Tráfico y Visitas de Usuarios
            </h3>
            <p className="text-xs text-gray-400 mt-0.5">
              Métricas geográficas y temporales anónimas de clientes visitando tu catálogo.
            </p>
          </div>

          {visitsQ.isLoading || !visits ? (
            <p className="text-gray-500 text-sm py-10 text-center">Cargando estadísticas de visitas...</p>
          ) : (
            <>
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
                <StatCard label="Total Histórico" value={visits.total} icon={<Eye size={16} />} />
                <StatCard label="Hoy" value={visits.today} icon={<Clock size={16} />} />
                <StatCard label="Últimos 7 días" value={visits.last7Days} icon={<Calendar size={16} />} />
                <StatCard label="Últimos 30 días" value={visits.last30Days} icon={<TrendingUp size={16} />} />
              </div>

              {peakVisitHour && peakVisitHour.count > 0 && (
                <div className="bg-[#161626] p-3.5 rounded-2xl border border-white/5 flex items-center justify-between text-xs uppercase tracking-wider text-gray-400">
                  <span>Hora pico de visitas (7 días):</span>
                  <span className="text-[#ffd025] font-black text-sm">
                    {String(peakVisitHour.hour).padStart(2, "0")}:00 · {peakVisitHour.count} visitas
                  </span>
                </div>
              )}

              <div className="bg-[#161626] p-4 rounded-2xl border border-white/5 space-y-3">
                <h4 className="text-xs uppercase font-bold text-gray-300 flex items-center gap-2">
                  <Clock size={14} className="text-[#ffd025]" /> Visitas por hora (últimos 7 días)
                </h4>
                <div className="flex items-end gap-1 h-32 bg-black/30 rounded-xl p-3 border border-white/5">
                  {visits.hourlyDistribution.map((b) => (
                    <div key={b.hour} className="flex-1 flex flex-col items-center gap-1 group">
                      <div className="text-[9px] text-gray-500 group-hover:text-[#ffd025]">
                        {b.count > 0 ? b.count : ""}
                      </div>
                      <div
                        className="w-full bg-gradient-to-t from-[#ffd025]/30 to-[#ffd025] rounded-t transition-all"
                        style={{ height: `${(b.count / maxVisitHour) * 85}%`, minHeight: b.count > 0 ? 4 : 0 }}
                        title={`${String(b.hour).padStart(2, "0")}:00 - ${b.count} visitas`}
                      />
                      <div className="text-[9px] text-gray-500">{b.hour}</div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="grid md:grid-cols-2 gap-4">
                <div className="bg-[#161626] p-4 rounded-2xl border border-white/5">
                  <h4 className="text-xs uppercase font-bold text-gray-300 mb-3 flex items-center gap-2">
                    <MapPin size={14} className="text-[#ffd025]" /> Top ciudades (30 días)
                  </h4>
                  {visits.topCities.length === 0 ? (
                    <p className="text-xs text-gray-500">Sin datos suficientes todavía.</p>
                  ) : (
                    <ul className="space-y-1.5">
                      {visits.topCities.map((c, i) => (
                        <li
                          key={`${c.city}-${i}`}
                          className="flex justify-between items-center text-xs bg-black/30 px-3 py-2 rounded-xl border border-white/5"
                        >
                          <span className="text-gray-200">
                            <span className="text-[#ffd025] font-black mr-2">#{i + 1}</span>
                            {c.city}
                            {c.country && <span className="text-gray-400 ml-1">· {c.country}</span>}
                          </span>
                          <span className="font-mono font-bold text-[#ffd025]">{c.count}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>

                <div className="bg-[#161626] p-4 rounded-2xl border border-white/5">
                  <h4 className="text-xs uppercase font-bold text-gray-300 mb-3 flex items-center gap-2">
                    <Globe size={14} className="text-[#ffd025]" /> Top países (30 días)
                  </h4>
                  {visits.topCountries.length === 0 ? (
                    <p className="text-xs text-gray-500">Sin datos suficientes todavía.</p>
                  ) : (
                    <ul className="space-y-1.5">
                      {visits.topCountries.map((c, i) => (
                        <li
                          key={`${c.country}-${i}`}
                          className="flex justify-between items-center text-xs bg-black/30 px-3 py-2 rounded-xl border border-white/5"
                        >
                          <span className="text-gray-200">
                            <span className="text-[#ffd025] font-black mr-2">#{i + 1}</span>
                            {c.country}
                          </span>
                          <span className="font-mono font-bold text-[#ffd025]">{c.count}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* Tab 3: Lo Más Vendido */}
      {statsSubTab === "top" && (
        <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 sm:p-8 rounded-3xl border border-white/10 shadow-2xl shadow-black/80 space-y-4">
          <div className="flex items-center justify-between pb-3 border-b border-white/10">
            <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
              <Flame size={18} className="text-orange-400" /> Ranking de Productos Estrella
            </h3>
            <span className="text-xs text-gray-400 font-mono">
              {sales?.topProducts.length ?? 0} artículos con ventas
            </span>
          </div>

          {!sales?.topProducts.length ? (
            <p className="text-gray-500 text-sm py-10 text-center">Aún no hay ventas registradas en este período.</p>
          ) : (
            <div className="space-y-3">
              {sales.topProducts.map((p, i) => (
                <div key={`${p.name}-${i}`} className="bg-[#161626] p-4 rounded-2xl border border-white/5 space-y-2">
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-white font-bold truncate pr-2 flex items-center gap-2">
                      <span className="text-[#ffd025] font-black text-xs">#{i + 1}</span>
                      <span>{p.name}</span>
                    </span>
                    <span className="font-black text-[#ffd025] shrink-0 font-mono">
                      {p.qty}u · ${p.revenue.toLocaleString("es-CL")}
                    </span>
                  </div>
                  <div className="h-2 bg-black/40 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-gradient-to-r from-[#ffd025] to-orange-500 rounded-full transition-all"
                      style={{ width: `${(p.qty / maxTopQty) * 100}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function CustomersAdminPanel() {
  const [customersData, setCustomersData] = useState<{
    registered: Array<{ id: number; email: string; name: string; phone: string; createdAt: string; orderCount: number; totalSpent: number }>;
    topBuyers: Array<{ name: string; phone: string; orderCount: number; totalSpent: number }>;
  } | null>(null);
  const [discounts, setDiscounts] = useState<Array<{
    id: number; code: string; type: string; amount: number; customerId: number | null;
    customerEmail: string | null; customerName: string | null; active: boolean;
    minOrder: number; usesLeft: number | null; expiresAt: string | null;
  }>>([]);
  const [loading, setLoading] = useState(true);
  const [customerTab, setCustomerTab] = useState<"top" | "directory" | "discounts">("top");
  const [customerSearch, setCustomerSearch] = useState("");
  const [newDiscount, setNewDiscount] = useState<{
    code: string; type: "percentage" | "fixed"; amount: number; customerId: number | null;
    minOrder: number; usesLeft: number | null; expiresAt: string;
  }>({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" });
  const [creating, setCreating] = useState(false);
  const [showCreateForm, setShowCreateForm] = useState(false);

  const loadData = async () => {
    setLoading(true);
    try {
      const [cRes, dRes] = await Promise.all([
        fetch("/api/admin/customers"),
        fetch("/api/admin/discounts"),
      ]);
      if (cRes.ok) setCustomersData(await cRes.json());
      if (dRes.ok) setDiscounts(await dRes.json());
    } catch {}
    setLoading(false);
  };

  useEffect(() => { loadData(); }, []);

  const createDiscount = async (e: FormEvent) => {
    e.preventDefault();
    if (!newDiscount.code.trim() || newDiscount.amount <= 0) return;
    setCreating(true);
    try {
      const r = await fetch("/api/admin/discounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...newDiscount, expiresAt: newDiscount.expiresAt || null }),
      });
      if (r.ok) {
        setNewDiscount({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" });
        setShowCreateForm(false);
        await loadData();
      }
    } catch {}
    setCreating(false);
  };

  const toggleDiscount = async (id: number, active: boolean) => {
    try {
      await fetch(`/api/admin/discounts/${id}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ active }) });
      setDiscounts(prev => prev.map(d => d.id === id ? { ...d, active } : d));
    } catch {}
  };

  const deleteDiscount = async (id: number) => {
    if (!confirm("¿Eliminar este descuento?")) return;
    try {
      await fetch(`/api/admin/discounts/${id}`, { method: "DELETE" });
      setDiscounts(prev => prev.filter(d => d.id !== id));
    } catch {}
  };

  if (loading) return (
    <div className="flex items-center justify-center py-20">
      <Loader2 size={32} className="animate-spin text-[#ffd025]" />
    </div>
  );

  const totalRegistered = customersData?.registered.length ?? 0;
  const topBuyersCount = customersData?.topBuyers.length ?? 0;
  const totalTopSpent = customersData?.topBuyers.reduce((acc, b) => acc + b.totalSpent, 0) ?? 0;
  const activeDiscountsCount = discounts.filter((d) => d.active).length;

  const filteredRegistered = (customersData?.registered ?? []).filter((c) => {
    if (!customerSearch.trim()) return true;
    const q = customerSearch.toLowerCase();
    return (
      c.name.toLowerCase().includes(q) ||
      c.email.toLowerCase().includes(q) ||
      (c.phone && c.phone.toLowerCase().includes(q))
    );
  });

  return (
    <div className="space-y-6 max-w-5xl mx-auto animate-fade-in">
      {/* Executive Header */}
      <div className="bg-[#13131f]/90 backdrop-blur-2xl p-6 rounded-3xl border border-white/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-xl shadow-black/60">
        <div>
          <div className="flex items-center gap-2 text-xs font-bold text-[#ffd025] uppercase tracking-wider mb-1">
            <Users size={15} /> CRM & Clientes
          </div>
          <h2 className="text-xl sm:text-2xl font-black text-white uppercase">Gestión de Clientes & Cupones</h2>
          <p className="text-xs text-gray-400 mt-1">
            Fideliza a tus compradores frecuentes, analiza consumo y crea promociones exclusivas.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => {
              setNewDiscount({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" });
              setShowCreateForm(!showCreateForm);
              setCustomerTab("discounts");
            }}
            className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-[#ffd025] to-[#ff9900] text-black font-black text-xs uppercase rounded-xl hover:scale-105 active:scale-95 transition-all shadow-md shadow-[#ffd025]/20"
          >
            <Plus size={15} />
            <span>Crear Cupón</span>
          </button>
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3.5">
        <div className="bg-[#13131f]/80 p-4 rounded-2xl border border-white/10 flex flex-col">
          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5">
            <Users size={12} className="text-[#ffd025]" /> Registrados
          </span>
          <span className="text-2xl font-black text-white mt-1">{totalRegistered}</span>
        </div>
        <div className="bg-[#13131f]/80 p-4 rounded-2xl border border-white/10 flex flex-col">
          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5">
            <TrendingUp size={12} className="text-[#ffd025]" /> Top Clientes
          </span>
          <span className="text-2xl font-black text-white mt-1">{topBuyersCount}</span>
        </div>
        <div className="bg-[#13131f]/80 p-4 rounded-2xl border border-white/10 flex flex-col">
          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5">
            <ShoppingBag size={12} className="text-[#ffd025]" /> Facturación Top
          </span>
          <span className="text-xl font-black text-[#ffd025] mt-1">${totalTopSpent.toLocaleString("es-CL")}</span>
        </div>
        <div className="bg-[#13131f]/80 p-4 rounded-2xl border border-white/10 flex flex-col">
          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5">
            <Gift size={12} className="text-[#ffd025]" /> Cupones Activos
          </span>
          <span className="text-2xl font-black text-emerald-400 mt-1">{activeDiscountsCount}</span>
        </div>
      </div>

      {/* Sub-Tabs Navigation */}
      <div className="flex items-center gap-2 border-b border-white/10 pb-3 overflow-x-auto scrollbar-hide">
        {[
          { id: "top", label: "1. Top Compradores", count: topBuyersCount, icon: TrendingUp },
          { id: "directory", label: "2. Directorio Clientes", count: totalRegistered, icon: Users },
          { id: "discounts", label: "3. Cupones & Descuentos", count: discounts.length, icon: Gift },
        ].map((tab) => {
          const Icon = tab.icon;
          const isActive = customerTab === tab.id;
          return (
            <button
              key={tab.id}
              onClick={() => setCustomerTab(tab.id as any)}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap border ${
                isActive
                  ? "bg-[#ffd025] text-black border-[#ffd025] shadow-lg shadow-[#ffd025]/20"
                  : "bg-white/5 text-gray-400 border-white/10 hover:text-white hover:bg-white/10"
              }`}
            >
              <Icon size={14} className={isActive ? "text-black" : "text-[#ffd025]"} />
              <span>{tab.label}</span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-mono ${isActive ? "bg-black/20 text-black font-black" : "bg-white/10 text-gray-400"}`}>
                {tab.count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Tab 1: Top Compradores */}
      {customerTab === "top" && (
        <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 p-6 shadow-2xl shadow-black/80 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
              <TrendingUp size={18} className="text-[#ffd025]" /> Ranking de Clientes Más Fieles
            </h3>
            <span className="text-xs text-gray-400">{topBuyersCount} compradores destacados</span>
          </div>

          {!customersData?.topBuyers.length ? (
            <p className="text-gray-500 text-sm py-10 text-center">Aún no hay compras registradas para calcular el ranking.</p>
          ) : (
            <div className="space-y-2.5">
              {customersData.topBuyers.map((b, i) => (
                <div key={i} className="flex items-center justify-between bg-[#161626] border border-white/5 hover:border-white/15 transition-all rounded-2xl px-4 py-3.5">
                  <div className="flex items-center gap-3">
                    <span className={`w-8 h-8 rounded-xl flex items-center justify-center font-black text-xs shrink-0 ${
                      i === 0 ? "bg-amber-400 text-black shadow-md shadow-amber-400/30" :
                      i === 1 ? "bg-slate-300 text-black" :
                      i === 2 ? "bg-amber-700 text-white" :
                      "bg-white/5 text-gray-400"
                    }`}>
                      #{i + 1}
                    </span>
                    <div>
                      <p className="font-bold text-sm text-white">{b.name || "Cliente Frecuente"}</p>
                      {b.phone && <p className="text-gray-400 text-xs font-mono">{b.phone}</p>}
                    </div>
                  </div>
                  <div className="text-right">
                    <p className="text-[#ffd025] font-black text-sm">${b.totalSpent.toLocaleString("es-CL")}</p>
                    <p className="text-gray-400 text-xs">{b.orderCount} {b.orderCount === 1 ? "pedido" : "pedidos"}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Tab 2: Directorio de Clientes */}
      {customerTab === "directory" && (
        <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 p-6 shadow-2xl shadow-black/80 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
              <Users size={18} className="text-[#ffd025]" /> Base de Datos de Clientes ({filteredRegistered.length})
            </h3>
            <input
              type="text"
              value={customerSearch}
              onChange={(e) => setCustomerSearch(e.target.value)}
              placeholder="Buscar por nombre, email o teléfono..."
              className="bg-black/40 border border-white/10 rounded-xl px-3.5 py-2 text-xs text-white placeholder-gray-500 focus:outline-none focus:border-[#ffd025] w-full sm:w-72"
            />
          </div>

          {filteredRegistered.length === 0 ? (
            <p className="text-gray-500 text-sm py-10 text-center">No se encontraron clientes registrados con esa búsqueda.</p>
          ) : (
            <div className="space-y-2.5">
              {filteredRegistered.map((c) => (
                <div key={c.id} className="bg-[#161626] border border-white/5 hover:border-white/15 rounded-2xl p-4 transition-all">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0 flex-1">
                      <p className="font-bold text-sm text-white truncate">{c.name}</p>
                      <p className="text-gray-400 text-xs truncate">{c.email}</p>
                      {c.phone && <p className="text-[#ffd025]/80 text-xs font-mono mt-0.5">{c.phone}</p>}
                    </div>
                    <div className="text-right shrink-0">
                      <p className="text-gray-500 text-[10px]">Registrado el {new Date(c.createdAt).toLocaleDateString("es-CL")}</p>
                      <button
                        onClick={() => {
                          setNewDiscount((p) => ({ ...p, customerId: c.id }));
                          setShowCreateForm(true);
                          setCustomerTab("discounts");
                        }}
                        className="text-xs text-[#ffd025] hover:underline mt-1 transition-colors font-bold flex items-center gap-1 ml-auto"
                      >
                        <Gift size={12} /> Dar cupón personal
                      </button>
                    </div>
                  </div>
                  <div className="flex items-center gap-3 mt-3 pt-2.5 border-t border-white/5 text-xs text-gray-400">
                    <span>
                      {c.orderCount === 0 ? (
                        <span className="text-gray-500">Sin pedidos aún</span>
                      ) : (
                        <><strong className="text-white">{c.orderCount}</strong> pedidos realizados</>
                      )}
                    </span>
                    {c.totalSpent > 0 && (
                      <>
                        <span>·</span>
                        <span>Total comprado: <strong className="text-[#ffd025]">${c.totalSpent.toLocaleString("es-CL")}</strong></span>
                      </>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Tab 3: Cupones y Descuentos */}
      {customerTab === "discounts" && (
        <div className="space-y-6">
          {showCreateForm && (
            <div className="bg-[#141422] rounded-3xl border border-[#ffd025]/30 p-6 shadow-2xl animate-fade-in space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-white/10">
                <h3 className="font-black text-sm uppercase text-[#ffd025] flex items-center gap-2">
                  <Plus size={16} /> Crear Código de Descuento
                </h3>
                <button
                  type="button"
                  onClick={() => setShowCreateForm(false)}
                  className="text-gray-400 hover:text-white"
                >
                  <X size={18} />
                </button>
              </div>

              <form onSubmit={createDiscount} className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3.5">
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Código del Cupón *</label>
                    <input
                      type="text"
                      value={newDiscount.code}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, code: e.target.value.toUpperCase() }))}
                      className="w-full bg-black/40 border border-white/10 rounded-xl p-2.5 text-white text-sm uppercase tracking-widest focus:outline-none focus:border-[#ffd025]"
                      placeholder="VERANO2026"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Tipo de Rebaja</label>
                    <select
                      value={newDiscount.type}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, type: e.target.value as "percentage" | "fixed" }))}
                      className="w-full bg-[#1b1b2d] border border-white/10 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                    >
                      <option value="percentage">% Porcentaje</option>
                      <option value="fixed">$ Monto Fijo</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">
                      {newDiscount.type === "percentage" ? "Porcentaje (%) *" : "Monto en Dinero ($) *"}
                    </label>
                    <input
                      type="number"
                      min={1}
                      max={newDiscount.type === "percentage" ? 100 : undefined}
                      value={newDiscount.amount || ""}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, amount: Number(e.target.value) }))}
                      className="w-full bg-black/40 border border-white/10 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                      placeholder="10"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Mínimo de Compra ($)</label>
                    <input
                      type="number"
                      min={0}
                      value={newDiscount.minOrder || ""}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, minOrder: Number(e.target.value) }))}
                      className="w-full bg-black/40 border border-white/10 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                      placeholder="0"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Límite de Usos (vacío = ilimitado)</label>
                    <input
                      type="number"
                      min={1}
                      value={newDiscount.usesLeft ?? ""}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, usesLeft: e.target.value ? Number(e.target.value) : null }))}
                      className="w-full bg-black/40 border border-white/10 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                      placeholder="Sin límite"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-gray-400 uppercase font-bold mb-1">Fecha de Expiración</label>
                    <input
                      type="date"
                      value={newDiscount.expiresAt}
                      onChange={(e) => setNewDiscount((p) => ({ ...p, expiresAt: e.target.value }))}
                      className="w-full bg-black/40 border border-white/10 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                    />
                  </div>
                </div>

                {newDiscount.customerId && (
                  <div className="flex items-center gap-2 p-2.5 rounded-xl bg-blue-500/10 border border-blue-500/20 text-xs text-blue-300">
                    <span>Asignado exclusivamente al cliente ID: <strong>{newDiscount.customerId}</strong></span>
                    <button
                      type="button"
                      onClick={() => setNewDiscount((p) => ({ ...p, customerId: null }))}
                      className="text-gray-400 hover:text-white ml-auto"
                    >
                      Quitar asignación
                    </button>
                  </div>
                )}

                <div className="flex gap-2 pt-2">
                  <button
                    type="submit"
                    disabled={creating}
                    className="flex-1 py-3 bg-[#ffd025] text-black font-black uppercase text-xs rounded-xl hover:scale-[1.01] active:scale-[0.99] transition-all disabled:opacity-50"
                  >
                    {creating ? "Guardando..." : "Guardar y Activar Cupón"}
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowCreateForm(false)}
                    className="px-4 py-3 bg-white/5 hover:bg-white/10 text-gray-400 rounded-xl text-xs font-bold transition-all"
                  >
                    Cancelar
                  </button>
                </div>
              </form>
            </div>
          )}

          <div className="bg-[#13131f]/90 backdrop-blur-2xl rounded-3xl border border-white/10 p-6 shadow-2xl space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-base font-black uppercase text-white flex items-center gap-2">
                <Gift size={18} className="text-[#ffd025]" /> Cupones Registrados ({discounts.length})
              </h3>
              {!showCreateForm && (
                <button
                  onClick={() => setShowCreateForm(true)}
                  className="text-xs text-[#ffd025] hover:underline font-bold flex items-center gap-1"
                >
                  <Plus size={14} /> Nuevo cupón
                </button>
              )}
            </div>

            {discounts.length === 0 ? (
              <p className="text-gray-500 text-sm py-8 text-center">No hay cupones creados aún. Haz clic en "Crear Cupón" para comenzar.</p>
            ) : (
              <div className="space-y-2.5">
                {discounts.map((d) => (
                  <div
                    key={d.id}
                    className={`flex items-center justify-between rounded-2xl p-4 border transition-all ${
                      d.active ? "bg-[#161626] border-white/10" : "bg-black/30 border-white/5 opacity-60"
                    }`}
                  >
                    <div className="flex-1 min-w-0 mr-3">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-black text-sm tracking-widest text-[#ffd025]">{d.code}</span>
                        {d.customerId && (
                          <span className="text-[10px] bg-blue-500/10 text-blue-300 px-2 py-0.5 rounded-full border border-blue-500/20 font-medium">
                            Personal
                          </span>
                        )}
                        {!d.active && (
                          <span className="text-[10px] bg-red-500/10 text-red-300 px-2 py-0.5 rounded-full border border-red-500/20 font-medium">
                            Pausado
                          </span>
                        )}
                      </div>
                      <p className="text-gray-400 text-xs mt-1 truncate">
                        {d.type === "percentage" ? `${d.amount}% de descuento` : `$${d.amount.toLocaleString("es-CL")} de descuento`}
                        {d.minOrder > 0 && ` · Mínimo $${d.minOrder.toLocaleString("es-CL")}`}
                        {d.usesLeft !== null && ` · ${d.usesLeft} usos restantes`}
                        {d.customerName && ` · Solo para ${d.customerName}`}
                        {d.expiresAt && ` · Vence ${new Date(d.expiresAt).toLocaleDateString("es-CL")}`}
                      </p>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        onClick={() => toggleDiscount(d.id, !d.active)}
                        className={`text-xs px-3 py-1.5 rounded-xl font-bold transition-all ${
                          d.active
                            ? "bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 border border-emerald-500/30"
                            : "bg-white/5 text-gray-400 hover:text-white border border-white/10"
                        }`}
                      >
                        {d.active ? "Activo" : "Pausado"}
                      </button>
                      <button
                        onClick={() => deleteDiscount(d.id)}
                        className="text-gray-500 hover:text-rose-400 p-2 rounded-xl hover:bg-rose-500/10 transition-colors"
                        title="Eliminar cupón"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({
  label,
  value,
  icon,
}: {
  label: string;
  value: number | string;
  icon: React.ReactNode;
}) {
  return (
    <div className="bg-[#141414] border border-gray-800 rounded-2xl p-4">
      <div className="flex items-center gap-2 text-gray-500 text-[10px] uppercase font-bold tracking-wider mb-1">
        <span className="text-[#ffd025]">{icon}</span>
        {label}
      </div>
      <div className="text-2xl font-black text-[#ffd025]">{value}</div>
    </div>
  );
}
