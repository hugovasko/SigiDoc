var map = L.map("map").setView([42.767285, 25.269495], 8);

L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution:
    '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}).addTo(map);

(async function () {
  try {
    const response = await fetch("map_points.json");
    const seals = await response.json();

    const markers = [];
    seals.forEach((seal) => {
      const [lat, lng] = seal.coordinates.split(", ");
      const title = seal.title || "No title";
      const date = seal.date || "No date";
      const filenameXml = seal.filename || null;
      const filenameHtml = filenameXml
        ? filenameXml.replace(".xml", ".html")
        : null;

      const pathToTheSeal = filenameHtml
        ? `http://127.0.0.1:9999/en/seals/${filenameHtml}`
        : null;
      markers.push(
        L.marker([lat, lng]).bindPopup(
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
