function doGet(e) {
  try {
    const action = (e.parameter.action || '').toString();

    if (action === 'downloadCountries') {
      const sheetName = e.parameter.sheet;
      if (!sheetName) throw new Error('Missing "sheet" parameter');

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) throw new Error(`Sheet "${sheetName}" is empty`);

      const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

      const csv = data.map(row => 
        row.map(cell => {
          const str = String(cell).replace(/"/g, '""');
          return `"${str}"`;
        }).join(',')
      ).join('\n');

      return ContentService
        .createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'GeoGuessr Stats API is running',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveGame') {
      const result = saveGameToSheets(data.gameData, data.userId);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Unknown action'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


function saveGameToSheets(gameData, userId) {
  try {
    const spreadsheetId = getOrCreateUserSpreadsheet(userId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    saveToMainSheet(spreadsheet, gameData);
    saveRoundsDetails(spreadsheet, gameData);
    saveCountryGuesses(spreadsheet, gameData);
    updateStatsSheet(spreadsheet, gameData);

    logEncounteredCountries(spreadsheet, gameData);

    return {
      success: true,
      message: 'Game saved successfully',
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function toModeLabelLower(gameData) {
  const m = (detectGameMode(gameData) || '').toString().toLowerCase();
  if (m === 'moving') return 'move';
  if (m === 'no move') return 'no move';
  if (m === 'nmpz') return 'nmpz';
  return 'custom';
}

function sanitizeSheetName(name) {
  return name.replace(/[\\\/\?\*\[\]\:]/g, ' ').trim().substring(0, 95);
}


function getOrCreateUserSpreadsheet(userId) {
  const folderName = 'GeoGuessr Stats Users';
  let folder;
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    folder = DriveApp.getRootFolder();
  }
  const fileName = `GeoGuessr Stats - ${userId}`;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) return files.next().getId();
  const spreadsheet = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (error) {}
  return spreadsheet.getId();
}

function detectGameMode(gameData) {
  if (gameData.restrictions.forbidMoving && !gameData.restrictions.forbidZooming && !gameData.restrictions.forbidRotating) {
    return 'No Move';
  }
  if (gameData.restrictions.forbidMoving && gameData.restrictions.forbidZooming && gameData.restrictions.forbidRotating) {
    return 'NMPZ';
  }
  if (!gameData.restrictions.forbidMoving && !gameData.restrictions.forbidZooming && !gameData.restrictions.forbidRotating) {
    return 'Moving';
  }
  const restrictions = [];
  if (gameData.restrictions.forbidMoving) restrictions.push('NM');
  if (gameData.restrictions.forbidZooming) restrictions.push('NZ');
  if (gameData.restrictions.forbidRotating) restrictions.push('NR');
  return restrictions.length > 0 ? restrictions.join('') : 'Custom';
}

function getCountryFromCoordinates(lat, lng) {
  try {
    const response = Maps.newGeocoder().reverseGeocode(lat, lng);
    if (response.results && response.results.length > 0) {
      const addressComponents = response.results[0].address_components;
      for (let component of addressComponents) {
        if (component.types.includes('country')) {
          return component.short_name.toLowerCase();
        }
      }
    }
  } catch (error) {
    console.log('Geocoding error:', error);
  }
  return 'unknown';
}

function saveToMainSheet(spreadsheet, gameData) {
  let sheet = spreadsheet.getSheetByName('Games');
  const maplink = (gameData.mapId) 
    ? `=HYPERLINK("https://www.geoguessr.com/maps/${gameData.mapId}";"Map Link")`
    : '';
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Games');
    const headers = [
      'Date', 'Token', 'Map Name', 'Map Link', 'Game Mode', 'Score', 'Distance (km)', 
      'Time Limit', 'Rounds', 'Perfect Score'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  const existingData = sheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][1] === gameData.token) return;
  }
  
  const row = [
    new Date(gameData.date),
    gameData.token,
    gameData.mapName,
    maplink,
    detectGameMode(gameData),
    gameData.score,
    Math.round(gameData.distance / 1000 * 100) / 100,
    gameData.timeLimit,
    gameData.rounds.length,
    gameData.score === 25000 ? 'YES' : 'NO'
  ];
  sheet.appendRow(row);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
         .sort({column: 1, ascending: false});
  }
}

function saveRoundsDetails(spreadsheet, gameData) {
  let sheet = spreadsheet.getSheetByName('Rounds');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Rounds');
    const headers = [
      'Date', 'Game Token', 'Map Name', 'Game Mode', 'Round', 
      'Score', 'Distance (km)', 'Actual Country', 'Actual Location'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  gameData.rounds.forEach(round => {
    const actualLocation = (round.lat && round.lng) ? 
      `=HYPERLINK("https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${round.lat},${round.lng}&heading=0&pitch=0";"Actual Location")`
      : '';




    const row = [
      new Date(gameData.date),
      gameData.token,
      gameData.mapName,
      detectGameMode(gameData),
      round.roundNumber,
      round.score,
      Math.round(round.distance / 1000 * 100) / 100,
      round.country.toUpperCase(),
      actualLocation
    ];

    sheet.appendRow(row);
  });
}

function saveCountryGuesses(spreadsheet, gameData) {
  let sheet = spreadsheet.getSheetByName('Country Recognition');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Country Recognition');
    const headers = [
      'Date', 'Game Token', 'Map Name', 'Game Mode', 'Round', 
      'Actual Country', 'Guessed Country', 'Correct Guess', 
      'Score', 'Distance (km)', 'Actual Location', 'Guess Location'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  gameData.rounds.forEach(round => {
    if (round.guessLat && round.guessLng) {
      const guessedCountry = getCountryFromCoordinates(round.guessLat, round.guessLng);
      const actualCountry = round.country || 'unknown';
      const correctGuess = guessedCountry === actualCountry ? 'YES' : 'NO';
      
      const actualLocation = (round.lat && round.lng) ? 
        `=HYPERLINK("https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${round.lat},${round.lng}&heading=0&pitch=0";"Actual Location")` 
        : '';
        
      const guessLocation = (round.guessLat && round.guessLng) ? 
        `=HYPERLINK("https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${round.guessLat},${round.guessLng}&heading=0&pitch=0";"Guess Location")` 
        : '';

      const row = [
        new Date(gameData.date),
        gameData.token,
        gameData.mapName,
        detectGameMode(gameData),
        round.roundNumber,
        actualCountry.toUpperCase(),
        guessedCountry.toUpperCase(),
        correctGuess,
        round.score,
        Math.round(round.distance / 1000 * 100) / 100,
        actualLocation,
        guessLocation
      ];
      sheet.appendRow(row);
    }
  });
}


function updateStatsSheet(spreadsheet, gameData) {
  let sheet = spreadsheet.getSheetByName('Statistics');
  if (!sheet) sheet = spreadsheet.insertSheet('Statistics');

  sheet.clear();
  initializeNewStatsSheet(sheet);

  const gamesSheet = spreadsheet.getSheetByName('Games');
  const countrySheet = spreadsheet.getSheetByName('Country Recognition');
  const allGames = gamesSheet.getDataRange().getValues().slice(1);
  const allCountryGuesses = countrySheet.getDataRange().getValues().slice(1);
  if (allGames.length === 0) return;

  const mapModeStats = {};

  allGames.forEach(game => {
    const mapName = game[2];
    const gameMode = game[4];
    const key = `${mapName}|||${gameMode}`;

    if (!mapModeStats[key]) {
      mapModeStats[key] = {
        mapName,
        gameMode,
        games: 0,
        totalScore: 0,
        bestScore: 0,
        worstScore: 25000,
        totalDistance: 0,
        perfectGames: 0,
        countryStats: {}
      };
    }

    mapModeStats[key].games++;
    mapModeStats[key].totalScore += game[5] || 0;
    mapModeStats[key].bestScore = Math.max(mapModeStats[key].bestScore, game[5] || 0);
    mapModeStats[key].worstScore = Math.min(mapModeStats[key].worstScore, game[5] || 0);
    mapModeStats[key].totalDistance += game[6] || 0;
    if (game[9] === 'YES') mapModeStats[key].perfectGames++;
  });

  allCountryGuesses.forEach(guess => {
    const mapName = guess[2];
    const gameMode = guess[3];
    const key = `${mapName}|||${gameMode}`;
    const actualCountry = guess[5];
    if (!actualCountry || actualCountry === 'UNKNOWN') return;

    if (!mapModeStats[key].countryStats[actualCountry]) {
      mapModeStats[key].countryStats[actualCountry] = 0;
    }
    mapModeStats[key].countryStats[actualCountry]++;
  });

  updateNewMapModeStatsSheet(sheet, mapModeStats);
}

function initializeNewStatsSheet(sheet) {
  sheet.getRange('A1').setValue('MAP + MODE STATISTICS').setFontWeight('bold').setFontSize(18);
  
  const headers = [
    'Map Name', 'Game Mode', 'Total Games', 'Avg Score', 'Best Score', 'Worst Score', 
    'Perfect Games', 'Perfect Rate %', 'Total Distance (km)', 'Avg Distance/Game (km)',
    'Countries Analysis', 'Most Difficult Country', 'Easiest Country', 'Total Countries Encountered'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(3);
}

function calculateAndUpdateNewMapModeStats(sheet, allGames, allCountryGuesses) {
  const mapModeStats = {};
  
  allGames.forEach(game => {
    const mapName = game[2];
    const gameMode = game[4];
    const key = `${mapName}|||${gameMode}`;
    
    if (!mapModeStats[key]) {
      mapModeStats[key] = { 
        mapName, 
        gameMode, 
        games: 0, 
        totalScore: 0, 
        bestScore: 0, 
        worstScore: 25000, 
        totalDistance: 0, 
        perfectGames: 0,
        countryStats: {}
      };
    }
    
    mapModeStats[key].games++;
    mapModeStats[key].totalScore += game[5] || 0;
    mapModeStats[key].bestScore = Math.max(mapModeStats[key].bestScore, game[5] || 0);
    mapModeStats[key].worstScore = Math.min(mapModeStats[key].worstScore, game[5] || 0);
    mapModeStats[key].totalDistance += game[6] || 0;
    if (game[9] === 'YES') mapModeStats[key].perfectGames++;
  });
  
  allCountryGuesses.forEach(guess => {
    const mapName = guess[2];
    const gameMode = guess[3];
    const key = `${mapName}|||${gameMode}`;
    const actualCountry = guess[5];
    const guessedCountry = guess[6];
    const correct = guess[7] === 'YES';
    
    if (mapModeStats[key] && actualCountry !== 'UNKNOWN') {
      if (!mapModeStats[key].countryStats[actualCountry]) {
        mapModeStats[key].countryStats[actualCountry] = {
          encountered: 0,
          correctGuesses: 0,
          totalScore: 0,
          wrongGuesses: {}
        };
      }
      
      mapModeStats[key].countryStats[actualCountry].encountered++;
      mapModeStats[key].countryStats[actualCountry].totalScore += guess[8] || 0;
      
      if (correct) {
        mapModeStats[key].countryStats[actualCountry].correctGuesses++;
      } else {
        if (!mapModeStats[key].countryStats[actualCountry].wrongGuesses[guessedCountry]) {
          mapModeStats[key].countryStats[actualCountry].wrongGuesses[guessedCountry] = 0;
        }
        mapModeStats[key].countryStats[actualCountry].wrongGuesses[guessedCountry]++;
      }
    }
  });
  
  updateNewMapModeStatsSheet(sheet, mapModeStats);
}

function updateNewMapModeStatsSheet(sheet, mapModeStats) {
  sheet.getRange('A4:Z10000').clearContent();
  
  let currentRow = 4;
  
  const sortedMapModes = Object.entries(mapModeStats).sort((a, b) => {
    if (a[1].mapName !== b[1].mapName) return a[1].mapName.localeCompare(b[1].mapName);
    return a[1].gameMode.localeCompare(b[1].gameMode);
  });
  
  sortedMapModes.forEach(([key, stats]) => {
    const perfectRate = stats.games > 0 ? Math.round((stats.perfectGames / stats.games) * 10000) / 100 : 0;
    const avgScore = stats.games > 0 ? Math.round(stats.totalScore / stats.games) : 0;
    const avgDistance = stats.games > 0 ? Math.round((stats.totalDistance / stats.games) * 100) / 100 : 0;
    
    let mostDifficultCountry = '';
    let easiestCountry = '';
    let lowestAccuracy = 100;
    let highestAccuracy = 0;
    const totalCountries = Object.keys(stats.countryStats).length;
    
    Object.entries(stats.countryStats).forEach(([country, countryData]) => {
      const accuracy = (countryData.correctGuesses / countryData.encountered) * 100;
      if (accuracy < lowestAccuracy) {
        lowestAccuracy = accuracy;
        mostDifficultCountry = `${country} (${Math.round(accuracy)}%)`;
      }
      if (accuracy > highestAccuracy) {
        highestAccuracy = accuracy;
        easiestCountry = `${country} (${Math.round(accuracy)}%)`;
      }
    });
    
    const mainRow = [
      stats.mapName,
      stats.gameMode,
      stats.games,
      avgScore,
      stats.bestScore,
      stats.worstScore === 25000 ? 0 : stats.worstScore,
      stats.perfectGames,
      perfectRate,
      Math.round(stats.totalDistance),
      avgDistance,
      totalCountries > 0 ? `${totalCountries} countries` : 'No data',
      mostDifficultCountry || 'N/A',
      easiestCountry || 'N/A',
      totalCountries
    ];
    
    sheet.getRange(currentRow, 1, 1, mainRow.length).setValues([mainRow]);
    
    if (Object.keys(stats.countryStats).length > 0) {
      let countryCol = 15;

      sheet.getRange(currentRow, countryCol, 1, 2).setValues([['Country', 'Encountered']]);
      sheet.getRange(currentRow, countryCol, 1, 2).setFontWeight('bold').setFontSize(9);
      let detailRow = currentRow + 1;

      const sortedCountries = Object.entries(stats.countryStats)
        .sort((a, b) => b[1].encountered - a[1].encountered);

      sortedCountries.forEach(([country, countryData]) => {
        sheet.getRange(detailRow, countryCol, 1, 2).setValues([[country, countryData.encountered]]);
        sheet.getRange(detailRow, countryCol, 1, 2).setFontSize(9);
        detailRow++;
      });
    }
    
    currentRow++;
  });
  
  sheet.autoResizeColumns(1, 20);
}


function logEncounteredCountries(spreadsheet, gameData) {
  const mapName = gameData.mapName || 'Unknown Map';
  const mode = toModeLabelLower(gameData);

  const sheetName = sanitizeSheetName(`${mapName} - ${mode}`);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const gamesSheet = spreadsheet.getSheetByName('Games');
  if (!gamesSheet) return;
  
  const allGames = gamesSheet.getDataRange().getValues().slice(1); // Skip header
  const filteredGames = allGames.filter(game => 
    game[2] === mapName && // Map Name
    toModeLabelLower({restrictions: getModeRestrictions(game[4])}) === mode // Game Mode
  );

  // Get ALL rounds for this map/mode combination from the Rounds sheet
  const roundsSheet = spreadsheet.getSheetByName('Rounds');
  if (!roundsSheet) return;
  
  const allRounds = roundsSheet.getDataRange().getValues().slice(1); // Skip header
  const filteredRounds = allRounds.filter(round => 
    round[2] === mapName && // Map Name
    toModeLabelLower({restrictions: getModeRestrictions(round[3])}) === mode // Game Mode
  );

  // Calculate general statistics
  let totalScore = 0;
  let totalDistance = 0;
  let gameCount = filteredGames.length;
  let perfectGames = 0;

  filteredGames.forEach(game => {
    totalScore += game[5] || 0; // Score column
    totalDistance += game[6] || 0; // Distance column
    if (game[9] === 'YES') perfectGames++; // Perfect Score column
  });

  const avgScore = gameCount > 0 ? Math.round(totalScore / gameCount) : 0;
  const avgDistance = gameCount > 0 ? Math.round((totalDistance / gameCount) * 100) / 100 : 0;
  const perfectRate = gameCount > 0 ? Math.round((perfectGames / gameCount) * 10000) / 100 : 0;

  // Main header
  sheet.getRange(1, 1).setValue(`${mapName} - ${mode.toUpperCase()}`);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);

  // General statistics on the right
  sheet.getRange(1, 4).setValue('GENERAL STATISTICS');
  sheet.getRange(1, 4).setFontWeight('bold').setFontSize(12);

  const generalStats = [
    ['Total Games:', gameCount],
    ['Average Score:', avgScore],
    ['Average Distance (km):', avgDistance],
    ['Perfect Games:', perfectGames],
    ['Perfect Rate (%):', perfectRate]
  ];

  sheet.getRange(2, 4, generalStats.length, 2).setValues(generalStats);
  sheet.getRange(2, 4, generalStats.length, 1).setFontWeight('bold');

  sheet.getRange(3, 1, 1, 2).setValues([['Country', 'Encountered']]);
  sheet.getRange(3, 1, 1, 2).setFontWeight('bold');
  sheet.setFrozenRows(3);

  const counts = {};
  filteredRounds.forEach(round => {
    const code = (round[7] || '').toString().toUpperCase(); // Actual Country column
    if (!code || code === 'UNKNOWN') return;
    counts[code] = (counts[code] || 0) + 1;
  });

  const rows = Object.entries(counts)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([code, n]) => [code, n]);

  if (rows.length) {
    sheet.getRange(4, 1, rows.length, 2).setValues(rows);
  }

  sheet.autoResizeColumns(1, 6);
}

// Fonction helper pour obtenir les restrictions à partir du nom du mode
function getModeRestrictions(gameMode) {
  if (gameMode === 'Moving') {
    return { forbidMoving: false, forbidZooming: false, forbidRotating: false };
  }
  if (gameMode === 'No Move') {
    return { forbidMoving: true, forbidZooming: false, forbidRotating: false };
  }
  if (gameMode === 'NMPZ') {
    return { forbidMoving: true, forbidZooming: true, forbidRotating: true };
  }
  return { forbidMoving: false, forbidZooming: false, forbidRotating: false }; // Custom default
}

function testSaveGame() {
  const testGameData = {
    token: 'test123',
    date: new Date().toISOString(),
    score: 24850,
    distance: 12500,
    mapName: 'World Test',
    mapId: 'world',
    gameMode: 'Moving',
    timeLimit: 0,
    rounds: [
      { 
        roundNumber: 1, score: 4970, distance: 2500, country: 'fr', 
        lat: 46.2276, lng: 2.2137, guessLat: 45.0, guessLng: 2.0 
      },
      { 
        roundNumber: 2, score: 4980, distance: 1500, country: 'de', 
        lat: 51.1657, lng: 10.4515, guessLat: 52.0, guessLng: 10.0 
      }
    ],
    restrictions: { forbidMoving: false, forbidZooming: false, forbidRotating: false }
  };
  const result = saveGameToSheets(testGameData);
  console.log(result);
}
