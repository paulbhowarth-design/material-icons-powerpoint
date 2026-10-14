
/* Material Symbols for PowerPoint — Mac edition */

const $ = id => document.getElementById(id);

let collection = null;
let names = [];
let filtered = [];
let visible = 80;

const suffixes = ['-rounded', '-sharp', '-outlined'];

function baseName(name) {
  for (const suffix of suffixes) {
    if (name.endsWith(suffix)) {
      return name.slice(0, -suffix.length);
    }
  }
  return name;
}

function selectedName(name) {
  const style = $('style').value;
  return style === 'outlined' ? name : `${name}-${style}`;
}

function getIcon(name, depth = 0) {
  if (depth > 8) return null;

  const icon = collection.icons?.[name];
  if (icon) return icon;

  const alias = collection.aliases?.[name];
  if (!alias) return null;

  const parent = getIcon(alias.parent, depth + 1);
  if (!parent) return null;

  return { ...parent, ...alias, body: parent.body };
}

function svgFor(name) {
  const icon = getIcon(selectedName(name)) || getIcon(name);

  if (!icon) {
    throw new Error('This icon style is unavailable');
  }

  const width = icon.width || collection.width || 24;
  const height = icon.height || collection.height || 24;
  const colour = $('color').value;

  let body = icon.body.replace(/currentColor/g, colour);

  const transforms = [];

  if (icon.hFlip) {
    transforms.push(`translate(${width} 0) scale(-1 1)`);
  }

  if (icon.vFlip) {
    transforms.push(`translate(0 ${height}) scale(1 -1)`);
  }

  if (icon.rotate) {
    transforms.push(
      `rotate(${icon.rotate * 90} ${width / 2} ${height / 2})`
    );
  }

  if (transforms.length) {
    body = `<g transform="${transforms.join(' ')}">${body}</g>`;
  }

  return `<svg xmlns="http://www.w3.org/2000/svg"
    width="96"
    height="96"
    viewBox="0 0 ${width} ${height}"
    fill="${colour}">
    ${body}
  </svg>`;
}

function render() {
  if (!collection) return;

  const query = $('q').value
    .trim()
    .toLowerCase()
    .replace(/[ _]+/g, '-');

  filtered = names.filter(name => name.includes(query));

  $('status').textContent =
    `${filtered.length.toLocaleString()} matching icons`;

  const fragment = document.createDocumentFragment();

  for (const name of filtered.slice(0, visible)) {
    const button = document.createElement('button');
    button.className = 'tile';
    button.title = `Insert ${name}`;

    try {
      const svg = svgFor(name);

      // Render SVG directly instead of as an image URL.
      const preview = document.createElement('span');
      preview.className = 'icon-preview';
      preview.innerHTML = svg;

      const svgElement = preview.querySelector('svg');
      svgElement.style.width = '32px';
      svgElement.style.height = '32px';
      svgElement.style.display = 'block';

      button.append(preview);
    } catch (error) {
      continue;
    }

    const label = document.createElement('span');
    label.textContent = name;

    button.append(label);
    button.addEventListener('click', () => insert(name));

    fragment.append(button);
  }

  $('grid').replaceChildren(fragment);
  $('more').hidden = visible >= filtered.length;
}

async function load() {
  try {
    const response = await fetch('./material-symbols.json', {
      cache: 'no-cache'
    });

    if (!response.ok) {
      throw new Error(
        `Material Symbols catalogue unavailable: HTTP ${response.status}`
      );
    }

    collection = await response.json();

    names = [
      ...new Set(
        [
          ...Object.keys(collection.icons || {}),
          ...Object.keys(collection.aliases || {})
        ].map(baseName)
      )
    ]
      .filter(name => getIcon(name))
      .sort();

    if (!names.length) {
      throw new Error('No icons found in catalogue');
    }

    render();
  } catch (error) {
    $('status').textContent = `Catalogue error: ${error.message}`;
  }
}

function encodeBase64(text) {
  const bytes = new TextEncoder().encode(text);
  let binary = '';

  for (let i = 0; i < bytes.length; i += 8192) {
    binary += String.fromCharCode(
      ...bytes.subarray(i, i + 8192)
    );
  }

  return btoa(binary);
}

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

    const svg = svgFor(name);
    const encoded = encodeBase64(svg);

    // Pass the base64 image data without the data URL prefix.
    const dataUrl = encoded;

    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        dataUrl,
        {
          coercionType: Office.CoercionType.Image,
          imageLeft: 90,
          imageTop: 90,
          imageWidth: 96,
          imageHeight: 96
        },
        result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(
              result.error?.message || 'PowerPoint rejected the SVG'
            ));
          }
        }
      );
    });

    $('status').textContent = `Inserted ${name}`;
  } catch (error) {
    $('status').textContent = `Insert failed: ${error.message}`;
    console.error(error);
  }
}

$('q').addEventListener('input', () => {
  visible = 80;
  render();
});

$('style').addEventListener('change', render);
$('color').addEventListener('input', render);

$('more').addEventListener('click', () => {
  visible += 80;
  render();
});

Office.onReady(load);
