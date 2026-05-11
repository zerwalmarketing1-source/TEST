# STOURS Catalogue Testing

Use this skill when testing the static STOURS DMC hotel catalogue frontend.

## Devin Secrets Needed

None. The catalogue is a public static frontend and does not require login or API credentials.

## Standard Runtime Test Flow

1. Open the public preview URL in Chrome.
2. Maximize or fullscreen the browser before recording GUI tests.
3. Start a screen recording for any browser-based UI validation.
4. Navigate from the hero CTA to the Marrakech section.
5. Verify the first Marrakech hotel card shows the reference data layout:
   - `ROOMS`
   - `RESTAURANTS`
   - `SIGNATURE EXPERIENCE`
   - `FACILITIES`
   - dashed key-information box below the data blocks
6. Open one gallery thumbnail and verify the dark lightbox overlay appears with an enlarged image.
7. Press Escape and verify the lightbox closes and the catalogue card is visible again.
8. Run a deployed-DOM audit in the browser console/CDP to verify every `article.hotel-card` has exactly four `.data-title` labels in this order: `ROOMS`, `RESTAURANTS`, `SIGNATURE EXPERIENCE`, `FACILITIES`.
9. Also check that no card displays `ACCESS Officiel`; `Officiel` should be treated as source metadata, not access/travel information.
10. Navigate to Casablanca and verify the visible header count matches the number of `article.hotel-card` elements in `#city-casablanca`.
11. Stop the recording and include the annotated recording in the final report.

## Useful DOM Snippets

All-card format audit:

```js
(() => {
  const expected = ['ROOMS', 'RESTAURANTS', 'SIGNATURE EXPERIENCE', 'FACILITIES'];
  const cards = [...document.querySelectorAll('article.hotel-card')];
  const failures = [];
  cards.forEach((card, index) => {
    const title = card.querySelector('h3')?.textContent.trim() || `card-${index + 1}`;
    const labels = [...card.querySelectorAll('.data-title')].map(el => el.textContent.trim());
    const access = [...card.querySelectorAll('.info small')]
      .find(el => el.textContent.trim() === 'ACCESS')
      ?.nextElementSibling?.textContent.trim() || '';
    if (labels.length !== 4 || labels.some((label, i) => label !== expected[i])) {
      failures.push({ index: index + 1, title, labels });
    }
    if (/ACCESS\s*Officiel/i.test(card.textContent) || /^Officiel$/i.test(access)) {
      failures.push({ index: index + 1, title, access });
    }
  });
  return { cardCount: cards.length, expectedLabels: expected, failureCount: failures.length, failures };
})();
```

Casablanca count audit:

```js
(() => {
  const section = document.querySelector('#city-casablanca');
  return {
    sectionFound: Boolean(section),
    heading: section?.querySelector('h2')?.textContent.trim(),
    headerCount: section?.querySelector('.city-count b')?.textContent.trim() || section?.querySelector('b')?.textContent.trim(),
    headerText: section?.querySelector('.city-count')?.textContent.trim() || '',
    cardCount: section ? section.querySelectorAll('article.hotel-card').length : 0
  };
})();
```

## Reporting

- Post one concise PR comment with runtime test results.
- Include the preview URL, annotated recording, screenshots, and DOM evidence.
- If any single assertion fails or is inconclusive, say testing was executed but do not declare success.
