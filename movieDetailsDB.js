function getMovieDetailsForEmptyCells() {
  const { apiKey } = require('./config');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(`A2:I${sheet.getLastRow()}`); // Assuming titles start from A2 and data goes up to column I
  const values = range.getValues();
  
  console.log(`Starting to process empty cells.`);
  console.log(`Total rows to process: ${values.length}`);

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2; // Adjust for 1-based indexing and header row
    const row = values[i];
    const originalTitle = row[0];
    const runtime = row[1];
    const releaseDate = row[2];
    const overview = row[3];
    const voteAverage = row[4];
    const genres = row[5];
    const coverArtUrl = row[6];
    const tmdbLink = row[7];
    const ageRating = row[8];
    
    if (originalTitle && (!runtime || !releaseDate || !overview || !voteAverage || !genres || !coverArtUrl || !tmdbLink || !ageRating)) {
      console.log(`Processing row: ${rowIndex}, Title: ${originalTitle}`);
      
      let searchData = searchMovie(originalTitle, apiKey);
      console.log(`Search Response for original title ${originalTitle}: ${JSON.stringify(searchData)}`);

      if (!searchData || searchData.results.length === 0) {
        const preprocessedTitle = preprocessTitle(originalTitle);
        searchData = searchMovie(preprocessedTitle, apiKey);
        console.log(`Search Response for preprocessed title ${preprocessedTitle}: ${JSON.stringify(searchData)}`);
      }

      if (searchData.results && searchData.results.length > 0) {
        const movie = searchData.results[0];
        const movieId = movie.id;
        const detailsUrl = `https://api.themoviedb.org/3/movie/${movieId}?api_key=${apiKey}&append_to_response=images`;
        const detailsResponse = UrlFetchApp.fetch(detailsUrl);
        const detailsData = JSON.parse(detailsResponse.getContentText());
        console.log(`Details Response for ${originalTitle}: ${JSON.stringify(detailsData)}`);

        if (detailsData.runtime && !runtime) {
          sheet.getRange(`B${rowIndex}`).setValue(detailsData.runtime); // Runtime in column B
          console.log(`Updated runtime for ${originalTitle}: ${detailsData.runtime}`);
        }

        if (detailsData.release_date && !releaseDate) {
          sheet.getRange(`C${rowIndex}`).setValue(detailsData.release_date); // Release date in column C
          console.log(`Updated release date for ${originalTitle}: ${detailsData.release_date}`);
        }

        if (detailsData.overview && !overview) {
          sheet.getRange(`D${rowIndex}`).setValue(detailsData.overview); // Overview in column D
          console.log(`Updated overview for ${originalTitle}: ${detailsData.overview}`);
        }

        if (detailsData.vote_average && !voteAverage) {
          sheet.getRange(`E${rowIndex}`).setValue(detailsData.vote_average); // Vote average in column E
          console.log(`Updated vote average for ${originalTitle}: ${detailsData.vote_average}`);
        }

        if (detailsData.genres && detailsData.genres.length > 0 && !genres) {
          const genresList = detailsData.genres.map(genre => genre.name).join(', ');
          sheet.getRange(`F${rowIndex}`).setValue(genresList); // Genres in column F
          console.log(`Updated genres for ${originalTitle}: ${genresList}`);
        }

        if (detailsData.poster_path && !coverArtUrl) {
          const coverArtUrl = `https://image.tmdb.org/t/p/w500${detailsData.poster_path}`;
          sheet.getRange(`G${rowIndex}`).setValue(coverArtUrl); // Cover art URL in column G
          console.log(`Updated cover art URL for ${originalTitle}: ${coverArtUrl}`);
        }

        if (!tmdbLink) {
          const tmdbLink = `https://www.themoviedb.org/movie/${movieId}`;
          sheet.getRange(`H${rowIndex}`).setValue(tmdbLink); // TMDB link in column H
          console.log(`Updated TMDB link for ${originalTitle}: ${tmdbLink}`);
        }

        // Fetch age rating
        if (!ageRating) {
          const ageRatingUrl = `https://api.themoviedb.org/3/movie/${movieId}/release_dates?api_key=${apiKey}`;
          const ageRatingResponse = UrlFetchApp.fetch(ageRatingUrl);
          const ageRatingData = JSON.parse(ageRatingResponse.getContentText());
          const ageRatingValue = getAgeRating(ageRatingData);
          if (ageRatingValue) {
            sheet.getRange(`I${rowIndex}`).setValue(ageRatingValue); // Age rating in column I
            console.log(`Updated age rating for ${originalTitle}: ${ageRatingValue}`);
          }
        }
      } else {
        console.log(`No results found for: ${originalTitle}`);
      }
    }
  }

  console.log(`Processing of empty cells completed.`);
}

function preprocessTitle(title) {
  title = title.replace(/\(.*?\)/g, '');
  title = title.replace(/[^a-zA-Z0-9 ]/g, '');
  title = title.trim();
  return title;
}

function searchMovie(title, apiKey) {
  const searchUrl = `https://api.themoviedb.org/3/search/movie?api_key=${apiKey}&query=${encodeURIComponent(title)}`;
  const searchResponse = UrlFetchApp.fetch(searchUrl);
  return JSON.parse(searchResponse.getContentText());
}

function getAgeRating(ageRatingData) {
  for (const result of ageRatingData.results) {
    if (result.iso_3166_1 === 'US') { 
      for (const release of result.release_dates) {
        if (release.certification) {
          return release.certification;
        }
      }
    }
  }
  return null;
}

function resetProgress() {
  PropertiesService.getScriptProperties().deleteProperty('lastProcessedIndex');
  console.log('Progress reset.');
}
