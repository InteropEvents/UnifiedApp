import { parseLocalDateKey } from './calendarDate';

describe('parseLocalDateKey', () => {
  it('keeps September 15 in the local calendar date', () => {
    const date = parseLocalDateKey('2026-09-15');

    expect(date.getFullYear()).toBe(2026);
    expect(date.getMonth()).toBe(8);
    expect(date.getDate()).toBe(15);
  });
});