app.post('/generate', async (req, res) => {
  const { slides } = req.body;

  if (!slides || !Array.isArray(slides)) {
    return res.status(400).json({ error: 'Missing or invalid slides array' });
  }

  const pres = new pptxgen();

  slides.forEach(slideData => {
    const slide = pres.addSlide();

    // Add the slide title
    slide.addText(slideData.title, {
      x: 0.5,
      y: 0.3,
      fontSize: 24,
      bold: true
    });

    // Flatten all text.value fields from objects[]
    const bodyText = slideData.objects
      .map(obj => obj.text?.value || '')
      .join('\n') // combine all text blocks
      .split('\n') // turn into array of bullet points
      .map(line => line.trim())
      .filter(line => line.length > 0); // remove empty lines

    // Add as bullets (if there's anything to show)
    if (bodyText.length > 0) {
      slide.addText(bodyText, {
        x: 0.7,
        y: 1.2,
        fontSize: 18,
        bullet: true,
        color: '363636',
        lineSpacingMultiple: 1.2
      });
    }
  });

  const buffer = await pres.stream();
  res.setHeader('Content-Disposition', 'attachment; filename="presentation.pptx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.send(buffer);
});
