# GSD Capture

GSD Capture is a phone-first voice capture assistant that sends tasks into Microsoft To Do, keeps spoken captures in a server-backed Review Inbox, and uses AI to suggest organization without silently making every decision for the user.

## Current Production features

- Private per-user Microsoft connection and Voice Capture key
- Apple Shortcut voice capture
- Server-backed GSD Review Inbox
- AI suggested folder, confidence, and explanation
- Approve, change, keep, complete, or delete review items
- Microsoft To Do reminders from explicit voice requests
- Microsoft Calendar events from explicit voice requests
- Calendar and reminder suggestions for review
- Friend beta onboarding checklist and troubleshooting tools

## Beta tester resources

- `beta-tester-guide.html` - mobile guide available inside the app
- `beta-tester-guide.pdf` - printable/shareable tester guide
- `BETA-LAUNCH-CHECKLIST.md` - Brian's pre-launch checklist

## Security notes

- Never commit API keys, Microsoft client secrets, Voice Capture keys, or KV tokens.
- The shared Apple Shortcut must contain a placeholder/import question, never Brian's personal Voice Key.
- Safe Support Details intentionally exclude the Voice Capture key.
