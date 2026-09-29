// Same range as the date slider on the Search page.
const DATE_MIN = 200;
const DATE_MAX = 1500;
const DATE_STEP = 20;

const TRANSLATIONS = {
  en: {
    date: "Date",
    suffix: "A.D.",
    reset: "Reset",
    seals: "seals",
    findspot: "Findspot",
    details: "View Details",
    noTitle: "No title",
    noDate: "No date",
    accuracy: {
      1: "High accuracy",
      2: "Medium accuracy",
      3: "Low accuracy",
      unknown: "Unknown accuracy",
      other: "Accuracy not specified",
    },
  },
  bg: {
    date: "Дата",
    suffix: "сл. Хр.",
    reset: "Изчисти",
    seals: "печата",
    findspot: "Местонамиране",
    details: "Виж детайли",
    noTitle: "Без заглавие",
    noDate: "Без дата",
    accuracy: {
      1: "Висока точност",
      2: "Средна точност",
      3: "Ниска точност",
      unknown: "Неизвестна точност",
      other: "Точността не е посочена",
    },
  },
};

const params = new URLSearchParams(window.location.search);
const language = /^[a-z]{2}$/.test(params.get("lang") || "")
  ? params.get("lang")
  : "en";
const t = TRANSLATIONS[language] || TRANSLATIONS.en;

var map = L.map("map", { scrollWheelZoom: false }).setView(
  [42.767285, 25.269495],
  7
);

// Zoom with the mouse wheel only after clicking the map, so that
// scrolling the page is not captured by the map.
map.on("focus", () => map.scrollWheelZoom.enable());
map.on("blur", () => map.scrollWheelZoom.disable());

L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution:
    '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}).addTo(map);

// Define custom icons for different findspot values
function coloredIcon(color) {
  return new L.Icon({
    iconUrl: `https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-${color}.png`,
    shadowUrl:
      "https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png",
    iconSize: [25, 41],
    iconAnchor: [12, 41],
    popupAnchor: [1, -34],
    shadowSize: [41, 41],
  });
}

var greenIcon = coloredIcon("green"); // Best findspot
var orangeIcon = coloredIcon("orange"); // Medium findspot
var redIcon = coloredIcon("red"); // Not good or unknown findspot

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function findspotIcon(findspot) {
  if (findspot === "1") return greenIcon;
  if (findspot === "2") return orangeIcon;
  return redIcon;
}

function findspotText(findspot) {
  if (t.accuracy[findspot]) return t.accuracy[findspot];
  if (findspot === "―") return t.accuracy.unknown;
  return t.accuracy.other;
}

// Seals without a usable date are always shown.
function isInRange(seal, from, to) {
  if (isNaN(seal.notBefore) || isNaN(seal.notAfter)) return true;
  return seal.notBefore <= to && seal.notAfter >= from;
}

(async function () {
  try {
    const response = await fetch("map_points.json");
    const points = await response.json();

    const seals = points.map((point) => {
      const [lat, lng] = point.coordinates.split(",").map((c) => c.trim());
      const filenameHtml = point.filename
        ? point.filename.replace(".xml", ".html")
        : null;

      const popupContent = `
        <div class="custom-popup">
          <div class="popup-title">${escapeHtml(point.title || t.noTitle)}</div>
          <div class="popup-date">${t.date}: ${escapeHtml(point.date || t.noDate)}</div>
          <div class="popup-findspot">${t.findspot}: ${findspotText(point.findspot)}</div>
          ${
            filenameHtml
              ? `<a href="/${language}/seals/${encodeURIComponent(filenameHtml)}" class="popup-button" target="_blank">${t.details}</a>`
              : ""
          }
        </div>
      `;

      return {
        notBefore: parseInt(point.notBefore, 10),
        notAfter: parseInt(point.notAfter, 10),
        marker: L.marker([lat, lng], {
          icon: findspotIcon(point.findspot),
        }).bindPopup(popupContent),
      };
    });

    const markersLayer = L.layerGroup().addTo(map);

    const fromInput = document.getElementById("date-from");
    const toInput = document.getElementById("date-to");
    const rangeLabel = document.getElementById("date-filter-range");
    const countLabel = document.getElementById("date-filter-count");
    const resetButton = document.getElementById("date-filter-reset");

    document.getElementById("date-filter-title").textContent = t.date;
    resetButton.textContent = t.reset;

    [fromInput, toInput].forEach((input) => {
      input.min = DATE_MIN;
      input.max = DATE_MAX;
      input.step = DATE_STEP;
    });

    function applyFilter() {
      const from = parseInt(fromInput.value, 10);
      const to = parseInt(toInput.value, 10);

      markersLayer.clearLayers();
      let visible = 0;
      seals.forEach((seal) => {
        if (isInRange(seal, from, to)) {
          markersLayer.addLayer(seal.marker);
          visible++;
        }
      });

      rangeLabel.textContent = `${from} – ${to} ${t.suffix}`;
      countLabel.textContent = `${visible} / ${seals.length} ${t.seals}`;
    }

    fromInput.addEventListener("input", () => {
      if (parseInt(fromInput.value, 10) > parseInt(toInput.value, 10)) {
        fromInput.value = toInput.value;
      }
      applyFilter();
    });

    toInput.addEventListener("input", () => {
      if (parseInt(toInput.value, 10) < parseInt(fromInput.value, 10)) {
        toInput.value = fromInput.value;
      }
      applyFilter();
    });

    resetButton.addEventListener("click", () => {
      fromInput.value = DATE_MIN;
      toInput.value = DATE_MAX;
      applyFilter();
    });

    fromInput.value = DATE_MIN;
    toInput.value = DATE_MAX;
    applyFilter();
  } catch (error) {
    console.error(error);
  }
})();
