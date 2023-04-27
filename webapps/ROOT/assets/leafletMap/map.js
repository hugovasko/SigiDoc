// Parse the JSON file to extract the coordinates
const coordinates = {
  "#Test1": "34.734564431148705, 32.48443089411125",
  "#Sozomenos": "35.067463372392375, 33.43608826042591",
};
const markers = [];
for (const [key, value] of Object.entries(coordinates)) {
  const [lat, lng] = value.split(", ");
  markers.push(L.marker([lat, lng]).bindPopup(key));
}

var map = L.map("map").setView([51.505, -0.09], 13);

L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution:
    '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}).addTo(map);

markers.forEach((marker) => marker.addTo(map));
