export const DEMO_MEETING_SUBJECT = 'Remote work strategy planning';

export const DEMO_MEETING_TRANSCRIPT = `WEBVTT

NOTE Synthetic transcript created for the UnifiedApp demo.

00:00:02.000 --> 00:00:16.000
Alex Morgan: Thanks for joining. Today we need to agree on a remote work pilot, the collaboration schedule, and how we will measure whether it is working.

00:00:18.000 --> 00:00:38.000
Priya Shah: The employee survey shows that 68 percent of respondents are more productive at home for focused work. The main concerns were slower onboarding and too many meetings across time zones.

00:00:40.000 --> 00:01:03.000
Jordan Lee: I propose a 90-day pilot from October 5 through December 18. Teams should use Tuesdays and Thursdays for live collaboration, with a two-hour overlap window selected by each team.

00:01:05.000 --> 00:01:26.000
Morgan Rivera: Security is comfortable with the pilot if everyone completes the device compliance check and managers confirm that confidential work is not done on shared computers.

00:01:28.000 --> 00:01:48.000
Alex Morgan: Let us make the pilot outcome-based. We will compare delivery predictability, employee sentiment, onboarding time, and after-hours meeting load against the September baseline.

00:01:50.000 --> 00:02:08.000
Priya Shah: I can draft the pilot policy and FAQ by September 18. I will include guidance for accessibility, caregivers, and employees who need office space.

00:02:10.000 --> 00:02:27.000
Jordan Lee: I will publish the manager toolkit by September 25 and run two manager office hours before the pilot starts.

00:02:29.000 --> 00:02:45.000
Morgan Rivera: I will validate the device and data-handling checklist with Security and Legal by September 22.

00:02:47.000 --> 00:03:05.000
Alex Morgan: I will create the pilot scorecard by September 28. Team leads will submit baseline data before October 2 and review results every two weeks.

00:03:07.000 --> 00:03:26.000
Priya Shah: We still need Finance to confirm whether the proposed 400-dollar home-office allowance can be included. That is an open item, not a final decision.

00:03:28.000 --> 00:03:47.000
Jordan Lee: The biggest risks are inconsistent manager expectations, isolation for new hires, and collaboration windows that disadvantage colleagues in other regions.

00:03:49.000 --> 00:04:05.000
Alex Morgan: Agreed. The pilot, Tuesday and Thursday collaboration pattern, and measurement plan are approved. We will revisit the allowance after Finance responds.`;

export function getDemoMeetingTranscript(subject: string): string | undefined {
  const normalizedSubject = subject.trim().toLocaleLowerCase();
  return normalizedSubject === DEMO_MEETING_SUBJECT.toLocaleLowerCase()
    ? DEMO_MEETING_TRANSCRIPT
    : undefined;
}