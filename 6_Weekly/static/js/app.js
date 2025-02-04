// Define the URL to fetch the data
let url = "http://127.0.0.1:5000/api/v1.0/team_monthly_plans";

const countryProfileMapping = {
    booster_doses_per_100people: "Booster doses per 100 people",
    country_name: "Country Name",
    gdp_2015: "GDP 2015",
    gdp_2016: "GDP 2016",
    gdp_2017: "GDP 2017",
    gdp_2018: "GDP 2018",
    gdp_2019: "GDP 2019",
    lat: "Latitude",
    lon: "Longitude",
    new_deaths: "New Deaths",
    newly_confirmed_cases: "Newly Confirmed Cases",
    newly_recovered_cases: "Newly Recovered Cases",
    total_confirmed_cases: "Total Confirmed Cases",
    total_deaths: "Total Deaths",
    total_recovered_cases: "Total Recovered Cases",
    total_vaccine_doses_administered_per_100population: "Total Vaccine Doses Administered Per 100 Population"
  };
function createCovidVaccineDoses(data, selectedCountryName) {
    // console.log("Creating stacked chart with most vaccinated per 100 population for country:", selectedCountryName);
    
      // Create an array of objects with country names and vaccine per 100 people
      let countriesWithVaccine = [];
      for (let countryName in data) {
          countriesWithVaccine.push({
              country_name: countryName,
              Vaccine_per100: data[countryName].total_vaccine_doses_administered_per_100population,
              
          });   
      }
      // Create an array of objects with country names and addtional dozes per 100 people
      let countriesWithDose = [];
      for (let countryName in data) {
          countriesWithDose.push({
              country_name: countryName,
              Dose_per100: data[countryName].booster_doses_per_100people,
              
          });
          
      }
      // Sort the array based on total vaccines per in descending order
    countriesWithVaccine.sort((a, b) => b.Vaccine_per100 - a.Vaccine_per100);
    // console.table(`This the mesagae ${countriesWithVaccine}`)
    // Extract data for the selected country
    let selectedCountryVaccineData = countriesWithVaccine.find(country => country.country_name === selectedCountryName);

     // Sort the array based on addtional vaccine doses per in descending order
     countriesWithDose.sort((a, b) => b.Dose_per100 - a.Dose_per100);
    //  console.table(`This the mesagae ${countriesWithVaccine}`)
     // Extract data for the selected country
     let selectedCountryDoseData = countriesWithVaccine.find(country => country.country_name === selectedCountryName);

    // If the selected country is found, include it in the top 10
    if (selectedCountryVaccineData) {
        countriesWithVaccine.unshift(selectedCountryVaccineData);
    }
    // If the selected country is found, include it in the top 10
    if (selectedCountryDoseData) {
        countriesWithDose.unshift(selectedCountryDoseData);
    }
    // Extract the label (country names) and values (vaccine per 100) for the top 10 
    let labels = countriesWithVaccine.slice(0, 11).map(country => country.country_name);
    let values = countriesWithVaccine.slice(0, 11).map(country => country.Vaccine_per100);
    
    // Extract the label (country names) and values (additional doses per 100) for the top 10 
    let labels2 = countriesWithDose.slice(0, 11).map(country => country.country_name);
    let values2 = countriesWithDose.slice(0, 11).map(country => country.Dose_per100);
    
    // Create the bar chart for vaccine per 100
    let BarData= {
         y: values,
         x: labels,
         name: 'Total Vaccine Doses Per 100 People',
        type: 'bar',
        marker: {
            color: 'red'
        }
    };

    // Create the bar chart for additional doses per 100
    let addtionalDose = {
        y: values2,
        x: labels2,
        name: 'Additional Does Per 100 People',
       type: 'bar',
       marker: {
        color: 'green'
    }
   };

   let layout1 = {
        title: `Top 10 Countries by Total Vaccine Doses<br> per 100 Population compare with ${selectedCountryName}`
    };
    let layout2 = {
        title: `Top 10 Countries by Additional Doses<br> per 100 Population compare with ${selectedCountryName}`
    };

    // Plot Chart
    Plotly.newPlot('bar-chart2', [BarData], layout1);
    Plotly.newPlot('bar-chart1', [addtionalDose], layout2);
   
}

 
function createlinechartGDP(data, selectedCountryName) {
    // Create an array of objects with country names, recovered cases, deaths, and GDP
    let GDP = [];
    let GDPRecovered = [];

    for (let countryName in data) {
        GDP.push({
            country_name: countryName,
            GDPall: ((data[countryName].gdp_2015 + data[countryName].gdp_2016 + data[countryName].gdp_2017 + data[countryName].gdp_2018 + data[countryName].gdp_2019))/5
        });        
    }

    // Sort the arrays based on total recovered cases and total deaths in descending order
    GDP.sort((a, b) => b.GDPall - a.GDPall);

    // Extract data for the selected country
    let selectedCountryGDP = GDP.find(country => country.country_name === selectedCountryName);


    // If the selected country is found, include it in the top 10
    if (selectedCountryGDP) {
        GDP.unshift(selectedCountryGDP);
    }
    
    // Extract labels (country names) and values for the top 10 or 11 countries
    let labels = GDP.slice(0, 10).map(country => country.country_name);
    let valuesGDP = GDP.slice(0, 10).map(country => country.GDPall);
    

    // Create the line charts
    let lineData = [{
        x: labels,
        y: valuesGDP,
        type: 'line',
        name: 'Average GDP Over 5 Years'
    }];

    let lineLayout = {
        title: `Top 10 countries with the highest average GDP over 5 years before Covid, Compared with ${selectedCountryName}`
    };

    // Plot Chart for total deaths and total recovered cases on the same line chart
    Plotly.newPlot('line-chart', lineData, lineLayout);
}

// Create a function to plot top 10 countries w/ most deaths
function createBarchartMostDeaths(data,selectedCountryName) {
    // Create an array of objects with country names and recovered cases
    // console.log(data);
    let countryObjects = [];
    // console.log(countryObjects);
    for (let countryName in data) {
        // console.log(countryName);

        countryObjects.push({
            country_name: countryName,
            total_deaths: data[countryName].total_deaths
        });
    }

    // 2. Filter data for plotting
    // Sort "countryObjects" in descending order using .sort() function based on the "Total Deaths" column
    let sortedArray = countryObjects.sort((a,b) => b.total_deaths - a.total_deaths);
    // console.log(sortedArray);

    // If the selected country is found, include it in the top 10
    let selectedCountryDeathsData = countryObjects.find(country => country.country_name === selectedCountryName);
    if (selectedCountryDeathsData) {
        countryObjects.unshift(selectedCountryDeathsData);
    }

    // Slice the first 10 countries/objects using .slice() function and save them in the array "top10Countries"
    let slicedArray = sortedArray.slice(0,11);
    // console.log(slicedArray);

    // Extract the country name for each country in "top10Countries" using .map() and arrow functions and save them in "namesTop10"    
    let namesTop10 = slicedArray.map(object => object.country_name);
    // console.log(namesTop10);

    // Extract the number of "Total Deaths" for each country in "top10Countries" using .map() and arrow functions and save them in "deathsTop10"
    let deathsTop10 = slicedArray.map(object => object.total_deaths)
    // console.log(deathsTop10);

    // 3. Plot the data
    let x_values = namesTop10;
    let y_values = deathsTop10;

    const plotData = [{
        x: x_values,
        y: y_values,
        type: 'bar'
    }];

    let layout = {
        title: `Top 10 Countries with the Most Deaths in COVID-19 Compared to ${selectedCountryName}`
    };
    // Update the bar chart
    Plotly.newPlot('bar', plotData, layout);
};



function createCountryMap(data, selectedCountryName) {
    // Define the map centered around a specific location (e.g., world map)
    let myMap = L.map("map", {
        center: [-21.977357, 80.239575],
        zoom: 3
    });

    // Add a tile layer to the map (you can choose a different tile layer)
    let street = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }).addTo(myMap);

    // Loop through the data and add markers for each country
    for (let countryName in data) {
        let country = data[countryName];
    
        // Check if the country has latitude and longitude information
        if (country.lat !== undefined && country.lon !== undefined) {
            // Convert lat and lon to numbers
            let latitude = Number(country.lat);
            let longitude = Number(country.lon);
    
            // Check if the conversion was successful
            if (!isNaN(latitude) && !isNaN(longitude)) {
                // Create a marker for the country
                let marker = L.marker([latitude, longitude]).addTo(myMap);
                
                // Add a popup with country information
                marker.bindPopup(`<b>${countryName}</b><br>${country.total_confirmed_cases} confirmed cases`);
    
                // // Highlight the selected country
                // if (countryName === selectedCountryName) {
                //     marker.openPopup();
                // }
            }
        }
    }
}


// Function to handle dropdown change event
function dropdownChange(data, selectedCountryName) {
    displaytwomonthsinfo(data, initialFieldTeamName);
    // displaysixweeksinfo(data, initialFieldTeamName);
}

// Function to display country info
function displaytwomonthsinfo(data, selectedCountryName) {
    // Find the selected country data using the provided country name
    let countryInfo = data[selectedCountryName];
    // console.log(countryInfo);
    // Clear any existing metadata
    let sampleMetadata = d3.select("#sample-metadata");
    sampleMetadata.html("");

    for (let [key, value] of Object.entries(countryInfo)) {
        sampleMetadata
            .append("p")
            .text(`${countryProfileMapping[key]}: ${value}`);   
    }
    
}

// Helper function to get the Monday of the current week
function getMonday(date) {
    date = new Date(date);
    let day = date.getDay(),
        diff = date.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(date.setDate(diff));
}

// Populate Dates dropdown based on the selection from the first dropdown
function datesDropdown(option) {
    let datesDropdown = d3.select("#seldatesBreakdown");
    datesDropdown.html(''); // Clear existing options

    if (option === "6 weeks") {
        let startDate = getMonday(new Date());
        for (let i = 0; i < 6; i++) {
            let mondayDate = new Date(startDate);
            mondayDate.setDate(startDate.getDate() + i * 7);
            let formattedDate = mondayDate.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'short', year: 'numeric' });
            let option = datesDropdown.append("option");
            option.property("value", formattedDate);
            option.text(formattedDate);
        }
    } else if (option === "2 months") {
        let currentDate = new Date();
        for (let i = 0; i < 2; i++) {
            let date = new Date(currentDate.getFullYear(), currentDate.getMonth() + i, 1);
            let formattedDate = date.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' });
            let option = datesDropdown.append("option");
            option.property("value", formattedDate);
            option.text(formattedDate);
        }
    }
}

// Initialize the page
function init() {
    // Fetch the JSON data from the above URL
    d3.json(url).then(data => {
        console.log("Fetched JSON data:", data);

        // Select the first dropdown element and populate it with "6 weeks" and "2 months"
        let planningtype = d3.select("#planningtype");
        ["2 months", "6 weeks"].forEach(optionText => {
            let option = planningtype.append("option");
            option.property("value", optionText);
            option.text(optionText);
        });

        // Extract the field team names from the data
        let fieldTeamNames = Object.keys(data);

        // Select the second dropdown element and populate it with field team names
        let fieldDropdown = d3.select("#selFieldTeamDropdown");
        fieldTeamNames.forEach(teamName => {
            let option = fieldDropdown.append("option");
            option.property("value", teamName);
            option.text(teamName);
        });

        // Define event for 6 weeks/2 months dropdown change and call the datesDropdown function
        planningtype.on("change", function () {
            let selectedOption = planningtype.property("value");
            datesDropdown(selectedOption);
        });

        // Define event for field team dropdown change and call the dropdownChange function
        fieldDropdown.on("change", function () {
            // Get the selected field team name
            let selectedFieldTeamName = fieldDropdown.property("value");
            dropdownChange(data, selectedFieldTeamName);
        });

        // Create the initial data and visualizations
        let initialFieldTeamName = fieldTeamNames[0];
        console.log("initialFieldTeamName:", fieldTeamNames[0]);
        displaytwomonthsinfo(data, initialFieldTeamName);
        datesDropdown("2 months")
        // displaysixweeksinfo(data, initialFieldTeamName);
    });
}

// Call the initialization function
init();