
async function insert(name) {
  $('status').textContent = `Inserting ${name}...`;

  try {
    if (
      typeof Office === 'undefined' ||
      !Office.context ||
      !Office.context.document
    ) {
      throw new Error('Open this add-in inside PowerPoint');
    }

    // Generate the SVG using the selected icon and colour.
    const svg = svgFor(name);

    // Insert genuine SVG content, not a PNG or image data URL.
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        svg,
        {
          coercionType: Office.CoercionType.XmlSvg,
          imageLeft: 90,
          imageTop: 90,
          imageWidth: 96,
          imageHeight: 96
        },
        result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(
              new Error(
                result.error?.message ||
                'PowerPoint rejected SVG insertion'
              )
            );
          }
        }
      );
    });

    $('status').textContent = `Inserted ${name} as SVG`;

  } catch (error) {
    $('status').textContent =
      `SVG insertion failed: ${error.message}`;

    console.error('SVG insertion error:', error);
  }
}
