var map = L.map("map").setView([42.767285, 25.269495], 8);

L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution:
    '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}).addTo(map);

(async function () {
  try {
    const response = await fetch("map_points.json");
    const coordinates = await response.json();

    const markers = [];
    for (const [key, value] of Object.entries(coordinates)) {
      const [lat, lng] = value.split(", ");
      markers.push(L.marker([lat, lng]).bindPopup(key));
    }

    markers.forEach((marker) => marker.addTo(map));
  } catch (error) {
    console.error(error);
  }
})();
