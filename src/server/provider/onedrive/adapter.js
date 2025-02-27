const querystring = require('node:querystring')

const isDrive = (item) => {
  if (item.remoteItem) {
    return !!item.remoteItem.driveType
  }

  return !!item.driveType
}

const isFolder = (item) => {
  if (item.remoteItem) {
    return !!item.remoteItem.folder || isDrive(item)
  }

  return !!item.folder || isDrive(item)
}

const getItemSize = (item) => {
  return item.size
}

const getItemThumbnailUrl = (item) => {
  return item.thumbnails && item.thumbnails[0] ? item.thumbnails[0].medium.url : null
}

const getItemIcon = (item) => {
  return isFolder(item) ? 'folder' : getItemThumbnailUrl(item)
}

const getItemSubList = (item) => {
  return item.value
}

const getItemName = (item) => {
  return item.name || ''
}

const getMimeType = (item) => {
  return item.file ? item.file.mimeType : null
}

const getItemId = (item) => {
  if (item.remoteItem) {
    return item.remoteItem.id
  }
  return item.id
}

const getItemRequestPath = (item) => {
  if (isDrive(item)) {
    return `root?driveId=${getItemId(item)}`
  }

  let query = `?driveId=${item.parentReference.driveId}`
  if (item.remoteItem) {
    query = `?driveId=${item.remoteItem.parentReference.driveId}`
  }
  return getItemId(item) + query
}

const getItemModifiedDate = (item) => {
  return item.lastModifiedDateTime
}

const getNextPagePath = (data) => {
  if (!data['@odata.nextLink']) {
    return null
  }

  const query = { cursor: querystring.parse(data['@odata.nextLink']).$skiptoken }
  return `?${querystring.stringify(query)}`
}

module.exports = (res, username, includeSharepoint) => {
  const data = { username, items: [] }
  const items = getItemSubList(res)
  items.forEach((item) => {
    data.items.push({
      isFolder: isFolder(item),
      icon: getItemIcon(item),
      name: getItemName(item),
      mimeType: getMimeType(item),
      id: getItemId(item),
      thumbnail: getItemThumbnailUrl(item),
      requestPath: getItemRequestPath(item),
      modifiedDate: getItemModifiedDate(item),
      size: getItemSize(item),
    })
  })
  if (includeSharepoint) {
    data.items.push({
      isFolder: true,
      icon: 'folder',
      name: 'SharePoint',
      id: 'root',
      requestPath: 'root?siteId=/',
    })
  }

  data.nextPagePath = getNextPagePath(res)

  return data
}
