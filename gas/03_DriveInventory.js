function collectInventory_(iterator, limit) {
  if (!Number.isInteger(limit) || limit < 1) throw new Error('Invalid inventory limit');
  const items = [];
  while (items.length < limit && iterator.hasNext()) {
    const file = iterator.next();
    items.push({
      id: file.getId(),
      name: file.getName(),
      mimeType: typeof file.getMimeType === 'function' ? file.getMimeType() : 'application/vnd.google-apps.folder',
      url: file.getUrl(),
      updatedAt: file.getLastUpdated().toISOString()
    });
  }
  return { items: items, truncated: iterator.hasNext() };
}
