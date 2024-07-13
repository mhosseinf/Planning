function createMap(layerGroups) {
    // Create the tile layer that will be the background of our map.
    let streetmap = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    });

    // Create a baseMaps object to hold the streetmap layer.
    let baseMaps = {
        "Street Map": streetmap
    };

    // Create an overlayMaps object to hold the rainStations, GSTNStations, IWSSStations, and LevelStations layers.
    let overlayMaps = {
        "Rain Stations": layerGroups[0],
        "GSTN Stations": layerGroups[1],
        "IWSS Stations": layerGroups[2],
        "Level Stations": layerGroups[3],
    };

    // Create the map object with options.
    let map = L.map("map-id", {
        center: [-31.9505, 115.8605],
        zoom: 11,
        layers: [streetmap, ...layerGroups]  // Spread the layerGroups array elements
    });

    // Create a layer control, and pass it baseMaps and overlayMaps. Add the layer control to the map.
    L.control.layers(baseMaps, overlayMaps, {
        collapsed: false
    }).addTo(map);
}
function createMarkers(data) {
    // Initialize arrays to hold rain station markers and GSTN markers.
    let rainMarkers = [];
    let GSTNMarkers = [];
    let IWSSMarkers = [];
    let levelMarkers = [];
    console.log("Fetched data:", data);
    // Get an array of station objects from the data.
    let stations = Object.values(data);
    console.log("Fetched stations:", stations);
    // Loop through the station objects.
    for (let i = 0; i < stations.length; i++) {
        let station = stations[i];
    
        // Check the StationId property to determine the type of marker.
        if (station.StationId === 'Rain_gauge') {
            // For Rain_gauge station, create a rain marker and bind a popup.
            let rainMarker = L.marker([station.lat, station.lon])
                .bindPopup(
                    `<h3>${station.address}</h3><h3>Site No: ${station.Sitno}</h3>`
                );
    
            // Add the rain marker to the rainMarkers array.
            rainMarkers.push(rainMarker);
        } else if (station.StationId === 'GSTN') {
            // For GSTN station, create a GSTN marker and bind a popup.
            let GSTNMarker = L.marker([station.lat, station.lon])
                .bindPopup(
                    `<h3>${station.address}</h3><h3>Site No: ${station.Sitno}</h3>`
                );
    
            // Add the GSTN marker to the GSTNMarkers array.
            GSTNMarkers.push(GSTNMarker);
        } else if (station.StationId === 'IWSS') {
            // For IWSS station, create an IWSS marker and bind a popup.
            let IWSSMarker = L.marker([station.lat, station.lon])
                .bindPopup(
                    `<h3>${station.address}</h3><h3>Site No: ${station.Sitno}</h3>`
                );
    
            // Add the IWSS marker to the IWSSMarkers array.
            IWSSMarkers.push(IWSSMarker);
        } else if (station.StationId === 'Level') {
            // For level station, create a level marker and bind a popup.
            let levelMarker = L.marker([station.lat, station.lon])
                .bindPopup(
                    `<h3>${station.address}</h3><h3>Site No: ${station.Sitno}</h3>`
                );
    
            // Add the level marker to the levelMarkers array.
            levelMarkers.push(levelMarker);
        }
    }
    

    // Create layer groups for rain markers and GSTN markers.
    let rainLayerGroup = L.layerGroup(rainMarkers);
    let GSTNLayerGroup = L.layerGroup(GSTNMarkers);
    let IWSSLayerGroup = L.layerGroup(IWSSMarkers);
    let levelLayerGroup = L.layerGroup(levelMarkers);

    // Call the createMap function with both layer groups.
    createMap([rainLayerGroup, GSTNLayerGroup,IWSSLayerGroup,levelLayerGroup]);
}

// Call the function to create rain markers and map
createMarkers(alldata);