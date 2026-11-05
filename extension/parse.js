(function (root, factory) {
  const api = factory(root);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  Object.assign(root, api);
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  const CATEGORY_TYPE = {
    nd: "Новостройка",
    ss: "Вторичка",
    new: "Новостройка",
    secondary: "Вторичка",
  };

  const TRANSACTION_TYPE = {
    sell: "Продажа",
    rent: "Аренда",
    short_rent: "Посуточно",
    long_rent: "Длительная аренда",
  };

  function normalizeText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/ё/g, "е")
      .replace(/\s+/g, " ")
      .trim();
  }

  function findPhrases(title, description, phrases) {
    const list = phrases || root.PHRASES || [];
    const text = normalizeText(`${title || ""}\n${description || ""}`);
    const found = [];
    for (const phrase of list) {
      const needle = normalizeText(phrase);
      if (needle && text.includes(needle)) found.push(phrase);
    }
    return found;
  }

  function looksLikeItem(entry) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return false;
    if (entry.type === "banner") return false;
    if (entry.id == null || entry.id === "") return false;
    if (entry.type && entry.type !== "item") return false;
    return true;
  }

  function extractItems(payload) {
    const groups = [payload && payload.items, payload && payload.catalog && payload.catalog.items];
    const items = [];
    const seen = new Set();
    for (const group of groups) {
      if (!Array.isArray(group)) continue;
      for (const entry of group) {
        if (!looksLikeItem(entry)) continue;
        const id = String(entry.id);
        if (seen.has(id)) continue;
        seen.add(id);
        items.push(entry);
      }
    }
    return items;
  }

  function absoluteUrl(urlPath) {
    if (!urlPath) return "";
    if (/^https?:\/\//i.test(urlPath)) return urlPath;
    return "https://www.avito.ru" + (urlPath.startsWith("/") ? urlPath : `/${urlPath}`);
  }

  function formatMetro(geo) {
    const refs = geo && Array.isArray(geo.geoReferences) ? geo.geoReferences : [];
    return refs
      .map((ref) => {
        if (!ref || typeof ref !== "object") return "";
        const name = ref.content || ref.name || "";
        const after = ref.after ? String(ref.after).trim() : "";
        return [name, after].filter(Boolean).join(" ");
      })
      .filter(Boolean)
      .join("; ");
  }

  function formatDeal(info) {
    if (!info || typeof info !== "object") return "";
    const category = CATEGORY_TYPE[info.categoryType] || info.categoryType || "";
    const transaction = TRANSACTION_TYPE[info.transactionType] || info.transactionType || "";
    return [category, transaction].filter(Boolean).join(", ");
  }

  function formatSeller(item) {
    const logo = item.userLogo && typeof item.userLogo === "object" ? item.userLogo : {};
    const direct = [logo.name, logo.title, logo.alt, logo.sellerName, item.sellerName].find(
      (value) => typeof value === "string" && value.trim()
    );
    if (direct) return direct.trim();
    const link = typeof logo.link === "string" ? logo.link : "";
    const brand = link.match(/\/brand\/([^/?#]+)/);
    if (!brand) return "";
    try {
      return decodeURIComponent(brand[1]).replace(/[_-]+/g, " ");
    } catch (error) {
      return brand[1].replace(/[_-]+/g, " ");
    }
  }

  function toRecord(item, searchUrl) {
    const price = item.priceDetailed || {};
    const photos = Number.isFinite(item.imagesCount)
      ? item.imagesCount
      : Array.isArray(item.images)
        ? item.images.length
        : "";
    return {
      id: String(item.id),
      url: absoluteUrl(item.urlPath),
      title: item.title || "",
      price: price.fullString || price.string || "",
      pricePerM2: item.normalizedPrice || "",
      city: (item.addressDetailed && item.addressDetailed.locationName) || (item.location && item.location.name) || "",
      address: (item.coords && item.coords.address_user) || "",
      metro: formatMetro(item.geo),
      development: item.developmentName || "",
      dealType: formatDeal(item.realtyInfo),
      seller: formatSeller(item),
      photos,
      description: item.description || "",
      searchUrl: searchUrl || "",
    };
  }

  function mergeRecord(previous, next, now) {
    return {
      ...next,
      savedAt: previous && previous.savedAt ? previous.savedAt : now,
      updatedAt: now,
    };
  }

  return {
    normalizeText,
    findPhrases,
    extractItems,
    toRecord,
    mergeRecord,
  };
});
