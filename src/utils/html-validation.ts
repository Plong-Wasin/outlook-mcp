export function findUnsupportedCss(html: string): string | null {
  if (/linear-gradient/i.test(html)) {
    return "❌ HTML body contains 'linear-gradient', which classic Outlook desktop (Word rendering engine) does not support. Remove it or replace it with a solid background-color instead.";
  }
  return null;
}
