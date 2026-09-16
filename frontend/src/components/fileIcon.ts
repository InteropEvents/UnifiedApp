export type FileIconKind = 'folder' | 'image' | 'video' | 'audio' | 'pdf' | 'word' | 'powerpoint' | 'excel' | 'text' | 'file';

export interface FileIconDetails {
  name?: string;
  file?: { mimeType?: string };
  folder?: unknown;
}

export function getFileIconKind(file: FileIconDetails): FileIconKind {
  if (file.folder) return 'folder';

  const mimeType = file.file?.mimeType?.toLowerCase() || '';
  const extension = file.name?.split('.').pop()?.toLowerCase() || '';

  if (mimeType.startsWith('image/') || ['bmp', 'gif', 'jpeg', 'jpg', 'png', 'svg', 'tif', 'tiff', 'webp'].includes(extension)) return 'image';
  if (mimeType.startsWith('video/') || ['avi', 'm4v', 'mkv', 'mov', 'mp4', 'webm'].includes(extension)) return 'video';
  if (mimeType.startsWith('audio/') || ['aac', 'flac', 'm4a', 'mp3', 'ogg', 'wav'].includes(extension)) return 'audio';
  if (mimeType.includes('pdf') || extension === 'pdf') return 'pdf';
  if (['doc', 'docx'].includes(extension)) return 'word';
  if (['ppt', 'pptx'].includes(extension)) return 'powerpoint';
  if (['csv', 'xls', 'xlsx'].includes(extension)) return 'excel';
  if (mimeType.startsWith('text/') || ['log', 'md', 'txt'].includes(extension)) return 'text';
  return 'file';
}

export function getFileExtension(file: FileIconDetails): string | undefined {
  const extension = file.name?.includes('.')
    ? file.name.split('.').pop()?.toLowerCase()
    : undefined;

  if (extension) return extension;

  const fallbackExtensions: Partial<Record<FileIconKind, string>> = {
    image: 'png',
    video: 'mp4',
    audio: 'mp3',
    pdf: 'pdf',
    word: 'docx',
    powerpoint: 'pptx',
    excel: 'xlsx',
    text: 'txt',
  };

  return fallbackExtensions[getFileIconKind(file)];
}