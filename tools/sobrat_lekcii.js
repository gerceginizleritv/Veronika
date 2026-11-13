/* Скрипт для консоли Chrome.
   Открыть https://rki.msu.ru/student/lectures (с открытой сессией),
   вставить этот код в консоль и нажать Enter.
   Соберёт текст всех лекций и вопросов к экзамену в один файл. */

(async () => {
  const toText = (el) => {
    el.querySelectorAll('script,style,nav,header,footer,noscript').forEach(e => e.remove());
    el.querySelectorAll('br').forEach(b => b.replaceWith('\n'));
    el.querySelectorAll('p,div,li,tr,h1,h2,h3,h4,h5,h6,td').forEach(e => e.append('\n'));
    return el.textContent
      .replace(/[ \t ]+/g, ' ')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  };

  // 1. Собираем ссылки на страницы курса
  const seen = new Set();
  const pages = [];
  document.querySelectorAll('a[href]').forEach(a => {
    const url = a.href;
    if (!/\/student\/(info|lecture)/i.test(url)) return;
    if (seen.has(url)) return;
    seen.add(url);
    pages.push({ url, title: a.textContent.trim().replace(/\s+/g, ' ') || '(без названия)' });
  });

  // Страница с вопросами к экзамену — на случай, если её нет в списке
  const exam = 'https://rki.msu.ru/student/info?id=1415';
  if (!seen.has(exam)) pages.unshift({ url: exam, title: 'Вопросы к итоговому экзамену' });

  console.log('Найдено страниц:', pages.length);
  pages.forEach((p, i) => console.log(`  ${i + 1}. ${p.title}`));
  if (!pages.length) {
    console.warn('Ссылок не найдено. Проверьте, что открыта страница со списком лекций.');
    return;
  }

  // 2. Скачиваем каждую
  let out = `Собрано со страницы: ${location.href}\nДата: ${new Date().toLocaleString('ru-RU')}\nСтраниц: ${pages.length}\n`;
  for (const [i, p] of pages.entries()) {
    console.log(`Загружаю (${i + 1}/${pages.length}): ${p.title}`);
    try {
      const html = await (await fetch(p.url, { credentials: 'include' })).text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const main = doc.querySelector('main, article, .content, #content, .page-content') || doc.body;
      out += `\n\n${'='.repeat(70)}\n${p.title}\n${p.url}\n${'='.repeat(70)}\n\n${toText(main)}\n`;
    } catch (e) {
      out += `\n\n${'='.repeat(70)}\n${p.title}\n${p.url}\nОШИБКА ЗАГРУЗКИ: ${e}\n`;
    }
    await new Promise(r => setTimeout(r, 300));
  }

  // 3. Сохраняем одним файлом
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([out], { type: 'text/plain;charset=utf-8' }));
  a.download = 'rki_lekcii.txt';
  a.click();
  console.log('Готово. Файл rki_lekcii.txt сохранён в «Загрузки».');
})();
