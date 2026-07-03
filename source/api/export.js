const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

// Source files live in /source/ at the repo root
const SOURCE_DIR = path.join(process.cwd(), 'source');

async function buildDeck(slideRequests) {
  // Load all unique source files needed
  const neededFiles = [...new Set(slideRequests.map(r => r.originalFile))];
  
  const sourceBuffers = {};
  for (const filename of neededFiles) {
    const filepath = path.join(SOURCE_DIR, filename);
    if (!fs.existsSync(filepath)) {
      throw new Error(`Source file not found: ${filename}`);
    }
    sourceBuffers[filename] = fs.readFileSync(filepath);
  }

  // Load all source zips
  const sourceZips = {};
  for (const [name, buf] of Object.entries(sourceBuffers)) {
    sourceZips[name] = await JSZip.loadAsync(buf);
  }

  // Build ordered slideNum -> xmlFile maps per source
  const slideMaps = {};
  for (const [name, zip] of Object.entries(sourceZips)) {
    const presXml = await zip.file('ppt/presentation.xml').async('text');
    const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
    const rIdToFile = {};
    for (const m of presRels.matchAll(/Id="(rId\d+)"[^>]*Target="slides\/(slide\d+\.xml)"/g)) {
      rIdToFile[m[1]] = m[2];
    }
    const ordered = [];
    for (const m of presXml.matchAll(/<p:sldId\b[^>]*\br:id="(rId\d+)"/g)) {
      if (rIdToFile[m[1]]) ordered.push(rIdToFile[m[1]]);
    }
    slideMaps[name] = ordered; // index 0 = slide 1
  }

  // Use first requested source as structural base (theme, master, layouts)
  const primarySrc = neededFiles[0];
  const primaryZip = sourceZips[primarySrc];
  const outZip = new JSZip();

  // Copy everything from primary EXCEPT its slides (we'll add our own)
  for (const [filePath, file] of Object.entries(primaryZip.files)) {
    if (file.dir) continue;
    if (filePath.match(/^ppt\/slides\//)) continue;
    const data = await file.async('nodebuffer');
    outZip.file(filePath, data);
  }

  // Copy each requested slide
  const newSlideFiles = [];
  const copiedPaths = new Set();

  for (let i = 0; i < slideRequests.length; i++) {
    const req = slideRequests[i];
    const srcZip = sourceZips[req.originalFile];
    const srcMap = slideMaps[req.originalFile];
    const srcSlideFile = srcMap[req.originalSlideNumber - 1];

    if (!srcSlideFile) {
      throw new Error(`Slide ${req.originalSlideNumber} not found in ${req.originalFile}`);
    }

    const newSlideName = `slide${i + 1}.xml`;

    // Copy slide xml
    const slideXml = await srcZip.file(`ppt/slides/${srcSlideFile}`).async('nodebuffer');
    outZip.file(`ppt/slides/${newSlideName}`, slideXml);
    newSlideFiles.push(newSlideName);

    // Copy slide rels + all referenced assets
    const relsPath = `ppt/slides/_rels/${srcSlideFile}.rels`;
    if (srcZip.file(relsPath)) {
      const relsXml = await srcZip.file(relsPath).async('text');
      outZip.file(`ppt/slides/_rels/${newSlideName}.rels`, relsXml);

      for (const m of relsXml.matchAll(/Target="([^"]+)"/g)) {
        const target = m[1];
        let fullPath;
        if (target.startsWith('../')) {
          fullPath = 'ppt/' + target.slice(3);
        } else if (target.startsWith('/')) {
          fullPath = target.slice(1);
        } else {
          fullPath = 'ppt/slides/' + target;
        }

        if (!copiedPaths.has(fullPath) && srcZip.file(fullPath) && !outZip.file(fullPath)) {
          const data = await srcZip.file(fullPath).async('nodebuffer');
          outZip.file(fullPath, data);
          copiedPaths.add(fullPath);

          // Also copy slideLayout rels
          if (fullPath.match(/slideLayouts\/slideLayout\d+\.xml$/)) {
            const layoutRelsPath = fullPath.replace('/slideLayouts/', '/slideLayouts/_rels/') + '.rels';
            if (!copiedPaths.has(layoutRelsPath) && srcZip.file(layoutRelsPath) && !outZip.file(layoutRelsPath)) {
              const lrData = await srcZip.file(layoutRelsPath).async('nodebuffer');
              outZip.file(layoutRelsPath, lrData);
              copiedPaths.add(layoutRelsPath);
            }
          }
        }
      }
    }
  }

  // Rewrite presentation.xml with new slide list
  let presXml = await primaryZip.file('ppt/presentation.xml').async('text');
  const newSldIdLst = newSlideFiles.map((f, i) =>
    `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`
  ).join('');
  presXml = presXml.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${newSldIdLst}</p:sldIdLst>`
  );
  outZip.file('ppt/presentation.xml', presXml);

  // Rewrite presentation.xml.rels
  let presRels = await primaryZip.file('ppt/_rels/presentation.xml.rels').async('text');
  presRels = presRels.replace(/<Relationship[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g, '');
  const newSlideRels = newSlideFiles.map((f, i) =>
    `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${f}"/>`
  ).join('');
  presRels = presRels.replace('</Relationships>', newSlideRels + '</Relationships>');
  outZip.file('ppt/_rels/presentation.xml.rels', presRels);

  // Rewrite [Content_Types].xml
  let contentTypes = await primaryZip.file('[Content_Types].xml').async('text');
  contentTypes = contentTypes.replace(/<Override PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');
  const newContentEntries = newSlideFiles.map(f =>
    `<Override PartName="/ppt/slides/${f}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join('');
  contentTypes = contentTypes.replace('</Types>', newContentEntries + '</Types>');
  outZip.file('[Content_Types].xml', contentTypes);

  return await outZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// Vercel serverless handler
module.exports = async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { name, slides } = req.body;

    if (!slides || !Array.isArray(slides) || slides.length === 0) {
      return res.status(400).json({ error: 'No slides provided' });
    }

    const pptxBuffer = await buildDeck(slides);

    const filename = (name || 'deck').replace(/[^a-z0-9_\- ]+/gi, '_') + '.pptx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', pptxBuffer.length);
    res.send(pptxBuffer);

  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: err.message });
  }
};
