// Hero TV-noise → YouTube video sequence
(function () {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;

  function startSequence() {
    var noise = document.querySelector('.hero-tv-noise');
    var videoBox = document.querySelector('.hero-tv-video');
    var logoLink = document.querySelector('.hero-logo-link');
    var audioCtrl = document.getElementById('audio-controls');
    var audioBtn = document.getElementById('audio-btn');
    var slider = document.getElementById('volume-slider');
    if (!noise || !videoBox) return;

    var noiseStartedAt = 0;
    var playingHasFired = false;
    var NOISE_START_DELAY = 2500;   // logo görünme süresi
    var NOISE_DURATION = 2000;       // noise'ın 2. saniyesinde fadeout başlar
    var FADE_BUFFER = 500;           // PLAYING noise sonrası geldiyse, elapsed + bu kadar bekle

    // YouTube iframe'i SAYFA AÇILIR AÇILMAZ oluştur — videoBox opacity:0
    // olduğu için görünmez ama iframe arka planda yüklenip oynamaya başlar.
    // Bu sayede noise söndüğünde initial siyah frame yerine asıl video görünür.
    var iframe = document.createElement('iframe');
    iframe.id = 'yt-iframe';
    iframe.src =
      'https://www.youtube.com/embed/Bp88d4ioe5s' +
      '?autoplay=1&mute=1&controls=0&modestbranding=1&rel=0' +
      '&playsinline=1' +
      '&iv_load_policy=3&disablekb=1&fs=0&enablejsapi=1';
    iframe.allow = 'autoplay; encrypted-media; picture-in-picture';
    iframe.referrerPolicy = 'strict-origin-when-cross-origin';
    iframe.setAttribute('frameborder', '0');
    iframe.setAttribute('allowfullscreen', '');
    videoBox.appendChild(iframe);

    var faded = false;
    var fadeNoise = function () {
      if (faded) return;
      faded = true;
      noise.classList.remove('active');
      videoBox.classList.add('active');
      if (logoLink) logoLink.classList.add('hidden');
      if (audioCtrl) audioCtrl.classList.add('show');
    };

    // Video bitmeden 1.5 saniye önce fade out: duration ve currentTime'ı poll et
    var nearEndFired = false;
    var ytDuration = 0;
    var ytCurrent = 0;
    var pollInterval = null;
    var END_LEAD_TIME = 1.5; // saniye — bu kadar erken fade başlat

    var endedReached = function () {
      if (nearEndFired) return;
      nearEndFired = true;
      if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }
      videoBox.classList.remove('active');
      if (logoLink) logoLink.classList.remove('hidden');
      if (audioCtrl) audioCtrl.classList.remove('show');
      var replay = document.getElementById('replay-btn');
      if (replay) replay.classList.add('show');
    };

    var startTimePolling = function () {
      if (pollInterval) return;
      pollInterval = setInterval(function () {
        if (!iframe.contentWindow) return;
        iframe.contentWindow.postMessage(JSON.stringify({
          event: 'command', func: 'getCurrentTime'
        }), 'https://www.youtube.com');
        iframe.contentWindow.postMessage(JSON.stringify({
          event: 'command', func: 'getDuration'
        }), 'https://www.youtube.com');
      }, 300);
    };

    // YouTube postMessage event listener
    window.addEventListener('message', function (ev) {
      if (ev.origin !== 'https://www.youtube.com') return;
      try {
        var d = typeof ev.data === 'string' ? JSON.parse(ev.data) : ev.data;
        if (!d) return;
        if (d.event === 'onStateChange') {
          if (d.info === 1) {
            // PLAYING
            playingHasFired = true;
            if (noiseStartedAt) {
              var elapsed = Date.now() - noiseStartedAt;
              setTimeout(fadeNoise, elapsed + FADE_BUFFER);
            }
            startTimePolling();
          } else if (d.info === 0) {
            // ENDED — yedek (genelde poll'dan önce yakalanır)
            endedReached();
          }
        } else if (d.event === 'infoDelivery' && d.info) {
          if (typeof d.info.currentTime === 'number') ytCurrent = d.info.currentTime;
          if (typeof d.info.duration === 'number' && d.info.duration > 0) ytDuration = d.info.duration;
          if (!nearEndFired && ytDuration > 0 && ytCurrent >= ytDuration - END_LEAD_TIME) {
            endedReached();
          }
        }
      } catch (err) {}
    });

    // iframe yüklenince state-change subscription gönder
    iframe.addEventListener('load', function () {
      iframe.contentWindow.postMessage(JSON.stringify({
        event: 'listening', id: 1
      }), 'https://www.youtube.com');
      iframe.contentWindow.postMessage(JSON.stringify({
        event: 'command', func: 'addEventListener', args: ['onStateChange']
      }), 'https://www.youtube.com');
    });

    // 3.5s sonra noise'u aç (logo görünüyor → noise üstünü kapatır)
    setTimeout(function () {
      noise.classList.add('active');
      noiseStartedAt = Date.now();
      // PLAYING noise'tan ÖNCE fired ettiyse, sabit NOISE_DURATION kadar göster
      if (playingHasFired) {
        setTimeout(fadeNoise, NOISE_DURATION);
      }
    }, NOISE_START_DELAY);

    // Hiç PLAYING event'i gelmezse defensive fallback
    setTimeout(fadeNoise, NOISE_START_DELAY + NOISE_DURATION + 5000);

    // Tekrar izle handler — closure içinde state'lere erişebilir
    document.addEventListener('click', function (e) {
      var btn = e.target.closest && e.target.closest('#replay-btn');
      if (!btn) return;
      btn.classList.remove('show');
      // Polling state'ini reset et
      nearEndFired = false;
      ytCurrent = 0;
      iframe.contentWindow.postMessage(JSON.stringify({
        event: 'command', func: 'seekTo', args: [0, true]
      }), 'https://www.youtube.com');
      iframe.contentWindow.postMessage(JSON.stringify({
        event: 'command', func: 'playVideo', args: []
      }), 'https://www.youtube.com');
      videoBox.classList.add('active');
      if (logoLink) logoLink.classList.add('hidden');
      if (audioCtrl) audioCtrl.classList.add('show');
      startTimePolling();
    });
  }

  // YouTube postMessage komutu
  function postYT(func, args) {
    var iframe = document.getElementById('yt-iframe');
    if (!iframe || !iframe.contentWindow) return;
    iframe.contentWindow.postMessage(JSON.stringify({
      event: 'command', func: func, args: args || []
    }), 'https://www.youtube.com');
  }

  // Mute toggle (slider değerinde unmute eder)
  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('#audio-btn');
    if (!btn) return;
    var slider = document.getElementById('volume-slider');
    if (btn.classList.contains('is-on')) {
      postYT('mute');
      btn.classList.remove('is-on');
    } else {
      postYT('unMute');
      postYT('setVolume', [parseInt(slider.value, 10) || 30]);
      btn.classList.add('is-on');
    }
  });

  // Volume slider — değişince setVolume + auto-mute/unmute
  document.addEventListener('input', function (e) {
    if (!e.target || e.target.id !== 'volume-slider') return;
    var btn = document.getElementById('audio-btn');
    var v = parseInt(e.target.value, 10) || 0;
    postYT('setVolume', [v]);
    if (v > 0 && btn && !btn.classList.contains('is-on')) {
      postYT('unMute');
      btn.classList.add('is-on');
    }
    if (v === 0 && btn && btn.classList.contains('is-on')) {
      postYT('mute');
      btn.classList.remove('is-on');
    }
  });

  if (document.readyState === 'complete') {
    startSequence();
  } else {
    window.addEventListener('load', startSequence);
  }
})();
