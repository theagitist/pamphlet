// Lightweight XML helper for navigating PPTX XML without a DOM library.
// PPTX XML is parsed via DOMParser-like approach using Node's built-in module.

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
};

// xmldom NodeList is not iterable with for...of; convert to array
function nodeList(nl) {
  if (!nl || !nl.length) return [];
  const arr = [];
  for (let i = 0; i < nl.length; i++) arr.push(nl[i]);
  return arr;
}

// Tag matcher: converts 'a:rPr' → full namespaced check
function matchTag(node, prefixedTag) {
  const [prefix, local] = prefixedTag.split(':');
  const ns = NS[prefix];
  return node.localName === local && node.namespaceURI === ns;
}

// Find first child matching prefixed tag
function child(node, prefixedTag) {
  if (!node) return null;
  for (const c of nodeList(node.childNodes)) {
    if (c.nodeType === 1 && matchTag(c, prefixedTag)) return c;
  }
  return null;
}

// Find all children matching prefixed tag
function children(node, prefixedTag) {
  if (!node) return [];
  const result = [];
  for (const c of nodeList(node.childNodes)) {
    if (c.nodeType === 1 && matchTag(c, prefixedTag)) result.push(c);
  }
  return result;
}

// Find first descendant matching prefixed tag (depth-first)
function descendant(node, prefixedTag) {
  if (!node) return null;
  for (const c of nodeList(node.childNodes)) {
    if (c.nodeType === 1) {
      if (matchTag(c, prefixedTag)) return c;
      const found = descendant(c, prefixedTag);
      if (found) return found;
    }
  }
  return null;
}

// Find all descendants matching prefixed tag
function descendants(node, prefixedTag) {
  if (!node) return [];
  const result = [];
  function walk(n) {
    for (const c of nodeList(n.childNodes)) {
      if (c.nodeType === 1) {
        if (matchTag(c, prefixedTag)) result.push(c);
        walk(c);
      }
    }
  }
  walk(node);
  return result;
}

// Get attribute value (non-namespaced)
function attr(node, name) {
  if (!node) return null;
  return node.getAttribute(name) || null;
}

// Get attribute value with namespace prefix (e.g., 'r:id')
function nsAttr(node, prefixedAttr) {
  if (!node) return null;
  const [prefix, local] = prefixedAttr.split(':');
  const ns = NS[prefix];
  return node.getAttributeNS(ns, local) || null;
}

// Get text content of a node (direct text only, not descendants)
function textContent(node) {
  if (!node) return '';
  let text = '';
  for (const c of nodeList(node.childNodes)) {
    if (c.nodeType === 3) text += c.nodeValue;
  }
  return text;
}

// Get all element children regardless of tag
function allChildren(node) {
  if (!node) return [];
  const result = [];
  for (const c of nodeList(node.childNodes)) {
    if (c.nodeType === 1) result.push(c);
  }
  return result;
}

export { NS, matchTag, child, children, descendant, descendants, attr, nsAttr, textContent, allChildren };
