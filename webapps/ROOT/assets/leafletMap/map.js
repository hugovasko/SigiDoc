var map = L.map("map").setView([42.767285, 25.269495], 8);

L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution:
    '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}).addTo(map);

// Define custom icons for different findspot values
var greenIcon = new L.Icon({
  iconUrl:
    "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png",
  shadowUrl:
    "https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png",
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41],
});

var orangeIcon = new L.Icon({
  iconUrl:
    "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-orange.png",
  shadowUrl:
    "https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png",
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41],
});

var redIcon = new L.Icon({
  iconUrl:
    "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-red.png",
  shadowUrl:
    "https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png",
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41],
});

// Add custom CSS for better popup styling with more specific selectors to override Leaflet defaults
const customPopupStyle = `
  <style>
    .custom-popup {
      font-family: Arial, sans-serif;
      line-height: 1.4;
    }
    .popup-title {
      font-weight: bold;
      font-size: 14px;
      margin-bottom: 5px;
      color: #333;
    }
    .popup-date {
      color: #666;
      margin-bottom: 5px;
    }
    .popup-findspot {
      color: #666;
      margin-bottom: 8px;
    }
    /* More specific selector to override Leaflet styles */
    .leaflet-popup-content .popup-button,
    .leaflet-container a.popup-button {
      display: inline-block;
      padding: 5px 10px;
      background-color: #2c6eb2;
      color: white !important; /* Use !important to ensure this style takes precedence */
      text-decoration: none;
      border-radius: 3px;
      font-size: 12px;
      text-align: center;
      margin-top: 5px;
    }
    .leaflet-popup-content .popup-button:hover,
    .leaflet-container a.popup-button:hover {
      background-color: #1c5293;
      color: white !important;
    }
  </style>
`;

// Insert custom CSS into document head
document.head.insertAdjacentHTML("beforeend", customPopupStyle);

(async function () {
  try {
    const response = await fetch("map_points.json");
    const seals = await response.json();
    console.log(seals);

    const markers = [];
    seals.forEach((seal) => {
      const [lat, lng] = seal.coordinates.split(", ");
      const title = seal.title || "No title";
      const date = seal.date || "No date";
      const findspot = seal.findspot;
      const filenameXml = seal.filename || null;
      const filenameHtml = filenameXml
        ? filenameXml.replace(".xml", ".html")
        : null;

      let pathToTheSeal = null;
      if (filenameHtml) {
        const urlObj = new URL(window.location.href);
        pathToTheSeal = `${urlObj.origin}/en/seals/${filenameHtml}`;
      }

      let markerIcon;
      if (findspot === "1") {
        markerIcon = greenIcon; // Best findspot
      } else if (findspot === "2") {
        markerIcon = orangeIcon; // Medium findspot
      } else if (findspot === "3") {
        markerIcon = redIcon; // Not good findspot
      } else if (findspot === "―") {
        markerIcon = redIcon; // Unknown findspot
      } else {
        markerIcon = redIcon;
      }

      // Get findspot accuracy text
      let findspotText = "";
      if (findspot === "1") {
        findspotText = "High accuracy";
      } else if (findspot === "2") {
        findspotText = "Medium accuracy";
      } else if (findspot === "3") {
        findspotText = "Low accuracy";
      } else if (findspot === "―") {
        findspotText = "Unknown accuracy";
      } else {
        findspotText = "Accuracy not specified";
      }

      // Create improved popup content with better styling and organization
      const popupContent = `
        <div class="custom-popup">
          <div class="popup-title">${title}</div>
          <div class="popup-date">Date: ${date}</div>
          <div class="popup-findspot">Findspot: ${findspotText}</div>
          ${
            filenameHtml
              ? `<a href="${pathToTheSeal}" class="popup-button" target="_blank">View Details</a>`
              : ""
          }
        </div>
      `;

      markers.push(
        L.marker([lat, lng], { icon: markerIcon }).bindPopup(popupContent)
      );
    });

    markers.forEach((marker) => marker.addTo(map));
  } catch (error) {
    console.error(error);
  }
})();
