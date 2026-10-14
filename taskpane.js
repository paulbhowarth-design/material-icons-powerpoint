
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


/* PowerPoint presentation theme colour picker */

const THEME_COLOURS = [
  ['Dark1', 'Dark 1'],
  ['Light1', 'Light 1'],
  ['Dark2', 'Dark 2'],
  ['Light2', 'Light 2'],
  ['Accent1', 'Accent 1'],
  ['Accent2', 'Accent 2'],
  ['Accent3', 'Accent 3'],
  ['Accent4', 'Accent 4'],
  ['Accent5', 'Accent 5'],
  ['Accent6', 'Accent 6']
];

async function refreshThemeColours() {
  const palette = $('theme-colours');
  const message = $('theme-message');

  palette.replaceChildren();
  message.textContent = 'Reading presentation colours...';

  if (
    typeof PowerPoint === 'undefined' ||
    !Office.context.requirements.isSetSupported(
      'PowerPointApi', '1.10'
    )
  ) {
    message.textContent =
      'Theme colour reading is unavailable in this PowerPoint version.';
    return;
  }

  try {
    const colours = await PowerPoint.run(async context => {
      const slides = context.presentation.getSelectedSlides();
      slides.load('items');
      await context.sync();

      if (!slides.items.length) {
        throw new Error('Select a slide first');
      }

      const scheme = slides.items[0].themeColorScheme;

      const results = THEME_COLOURS.map(([key, label]) => ({
        key,
        label,
        result: scheme.getThemeColor(key)
      }));

      await context.sync();

      return results.map(item => ({
        label: item.label,
        hex: '#' + item.result.value.replace(/^#/, '')
      }));
    });

    for (const colour of colours) {
      const swatch = document.createElement('button');

      swatch.type = 'button';
      swatch.className = 'theme-swatch';
      swatch.style.backgroundColor = colour.hex;
      swatch.title = `${colour.label}: ${colour.hex}`;
      swatch.setAttribute('aria-label', swatch.title);

      swatch.addEventListener('click', () => {
        $('color').value = colour.hex;

        palette.querySelectorAll('.theme-swatch')
          .forEach(button => button.classList.remove('active'));

        swatch.classList.add('active');

        // Trigger the existing SVG preview update.
        $('color').dispatchEvent(
          new Event('input', { bubbles: true })
        );
      });

      palette.appendChild(swatch);
    }

    message.textContent =
      'Colours from the selected slide theme';

  } catch (error) {
    message.textContent =
      'Could not read theme colours: ' + error.message;

    console.error('Theme colour error:', error);
  }
}

$('refresh-theme').addEventListener(
  'click',
  refreshThemeColours
);

Office.onReady(() => {
  refreshThemeColours();
});

