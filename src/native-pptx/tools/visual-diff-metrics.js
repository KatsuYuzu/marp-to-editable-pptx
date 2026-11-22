const { PNG } = require('pngjs')

function resizePng(image, width, height) {
  if (image.width === width && image.height === height) return image

  const resized = new PNG({ width, height })
  const xScale = image.width / width
  const yScale = image.height / height

  for (let y = 0; y < height; y++) {
    const sourceY = Math.min(image.height - 1, Math.floor((y + 0.5) * yScale))
    for (let x = 0; x < width; x++) {
      const sourceX = Math.min(image.width - 1, Math.floor((x + 0.5) * xScale))
      const source = (sourceY * image.width + sourceX) * 4
      const target = (y * width + x) * 4
      resized.data.set(image.data.subarray(source, source + 4), target)
    }
  }

  return resized
}

function normalizePngPair(reference, comparison) {
  return {
    reference,
    comparison: resizePng(comparison, reference.width, reference.height),
  }
}

function computeRegionalDiff(
  data1,
  data2,
  width,
  height,
  {
    columns = 16,
    rows = 9,
    pixelThreshold = 0.18,
  } = {},
) {
  const threshold = Math.round(pixelThreshold * 255)
  const regions = []
  let wholeDifferent = 0

  for (let row = 0; row < rows; row++) {
    const y0 = Math.floor((row * height) / rows)
    const y1 = Math.floor(((row + 1) * height) / rows)
    for (let column = 0; column < columns; column++) {
      const x0 = Math.floor((column * width) / columns)
      const x1 = Math.floor(((column + 1) * width) / columns)
      let different = 0

      for (let y = y0; y < y1; y++) {
        for (let x = x0; x < x1; x++) {
          const offset = (y * width + x) * 4
          const delta =
            Math.max(
              Math.abs(data1[offset] - data2[offset]),
              Math.abs(data1[offset + 1] - data2[offset + 1]),
              Math.abs(data1[offset + 2] - data2[offset + 2]),
              Math.abs(data1[offset + 3] - data2[offset + 3]),
            )
          if (delta > threshold) different++
        }
      }

      const pixels = Math.max(1, (x1 - x0) * (y1 - y0))
      wholeDifferent += different
      regions.push({
        column,
        row,
        x: x0,
        y: y0,
        width: x1 - x0,
        height: y1 - y0,
        diffPct: different / pixels,
      })
    }
  }

  regions.sort(
    (left, right) =>
      right.diffPct - left.diffPct ||
      left.row - right.row ||
      left.column - right.column,
  )
  return {
    wholeSlidePct: wholeDifferent / (width * height),
    maxRegionPct: regions[0]?.diffPct ?? 0,
    regions,
  }
}

module.exports = { computeRegionalDiff, normalizePngPair, resizePng }
