function getArtwork() {
  // Filtrar por obras de arte moderno (aproximadamente desde 1850 hasta 1970)
  const url = 'https://api.artic.edu/api/v1/artworks?' + 
    'fields=id,title,image_id,date_display,artist_display,medium_display,style_title' +
    '&query[term][style_title]=technology' +
    '&limit=100';
  
  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    if (!json.data || json.data.length === 0) {
      throw new Error('No modern artwork data received');
    }

    // Seleccionar una obra al azar
    const artworks = json.data;
    const randomArtwork = artworks[Math.floor(Math.random() * artworks.length)];

    if (!randomArtwork.image_id) {
      throw new Error('Selected artwork has no image');
    }

    // Devolver datos más detallados de la obra
    return {
      title: randomArtwork.title,
      artist: randomArtwork.artist_display,
      date: randomArtwork.date_display,
      medium: randomArtwork.medium_display,
      style: randomArtwork.style_title,
      imageUrl: `https://www.artic.edu/iiif/2/${randomArtwork.image_id}/full/843,/0/default.jpg`
    };
  } catch (error) {
    console.error('Error fetching modern artwork:', error);
    return null;
  }
}
