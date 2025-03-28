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

// var grayIcon = new L.Icon({
//   iconUrl:
//     "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-grey.png",
//   shadowUrl:
//     "https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png",
//   iconSize: [25, 41],
//   iconAnchor: [12, 41],
//   popupAnchor: [1, -34],
//   shadowSize: [41, 41],
// });

(async function () {
  try {
    const response = await fetch("map_points.json");
    const seals = await response.json();

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

      const pathToTheSeal = filenameHtml
        ? `http://127.0.0.1:9999/en/seals/${filenameHtml}`
        : null;

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

      markers.push(
        L.marker([lat, lng], { icon: markerIcon }).bindPopup(
          `${title}<br>${date}${
            filenameHtml
              ? `<br><a href="${pathToTheSeal}" target="_blank">See more</a>`
              : ""
          }`
        )
      );
    });

    markers.forEach((marker) => marker.addTo(map));
  } catch (error) {
    console.error(error);
  }
})();
