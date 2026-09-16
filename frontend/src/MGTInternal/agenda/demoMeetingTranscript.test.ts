import {
  DEMO_MEETING_TRANSCRIPT,
  getDemoMeetingTranscript,
} from './demoMeetingTranscript';

describe('getDemoMeetingTranscript', () => {
  it('returns the synthetic transcript for the demo meeting', () => {
    expect(getDemoMeetingTranscript(' Remote work strategy planning ')).toBe(
      DEMO_MEETING_TRANSCRIPT
    );
  });

  it('matches the meeting subject without case sensitivity', () => {
    expect(getDemoMeetingTranscript('REMOTE WORK STRATEGY PLANNING')).toBeDefined();
  });

  it('does not attach the transcript to other meetings', () => {
    expect(getDemoMeetingTranscript('Quarterly business review')).toBeUndefined();
  });
});