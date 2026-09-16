import { getFileExtension, getFileIconKind } from './fileIcon';

describe('getFileIconKind', () => {
  it.each([
    [{ name: 'PowerApp-Issue.mp4', file: { mimeType: 'video/mp4' } }, 'video'],
    [{ name: 'Scala.txt', file: { mimeType: 'text/plain' } }, 'text'],
    [{ name: 'Info Demo ClickThroughs.pptx', file: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' } }, 'powerpoint'],
    [{ name: 'RECORDING.MP4', file: { mimeType: 'application/octet-stream' } }, 'video'],
    [{ name: 'NOTES.TXT', file: { mimeType: 'application/octet-stream' } }, 'text'],
  ])('maps %s to the %s icon', (file, expectedIcon) => {
    expect(getFileIconKind(file)).toBe(expectedIcon);
  });

  it('passes the normalized PowerPoint extension to Fluent file icons', () => {
    expect(getFileExtension({ name: 'Info Demo ClickThroughs.PPTX' })).toBe('pptx');
  });
});