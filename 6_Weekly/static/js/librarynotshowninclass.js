// Create the pie chart using Highcarts library
function createpichartWithMostRecovered(data, selectedCountryName) {
    // console.log("Creating pie chart with most recovered cases for country:", selectedCountryName);

    // Create an array of objects with country names and recovered cases
    let countriesWithRecovered = [];
    for (let countryName in data) {
        countriesWithRecovered.push({
            name: countryName,
            y: data[countryName].total_recovered_cases
        });
    }

    // Sort the array based on total recovered cases in descending order
    countriesWithRecovered.sort((a, b) => b.y - a.y);

    // Extract data for the selected country
    let selectedCountryRecoveredData = countriesWithRecovered.find(country => country.name === selectedCountryName);

    // If the selected country is found, include it in the top 10
    if (selectedCountryRecoveredData) {
        countriesWithRecovered.unshift(selectedCountryRecoveredData);
    }

    // Extract data for the top 10 or 11 countries
    let topCountries = countriesWithRecovered.slice(0, 11);

    // Create the pie chart
    Highcharts.chart('pie', {
        chart: {
            type: 'pie'
        },
        plotOptions: {
            pie: {
                dataLabels: {
                    style: {
                        fontSize: '12px'
                    }
                }
            }
        },
        title: {
            text: `Top 10 Country Recovered Cases Comparison with ${selectedCountryName}`
        },
        series: [{
            name: 'Recovered Cases',
            data: topCountries
        }]
    });
}