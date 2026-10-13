#!/usr/bin/env python3
import asyncio, json, os, shutil, subprocess, time, glob, signal
import websocket

SHOTS = '/home/geeker/architecture/home/geeker/code_repo/best_practice_perf/labs_ai_/agent_app/intro_self/slides'
os.makedirs(SHOTS, exist_ok=True)
for f in glob.glob(os.path.join(SHOTS, '*.png')):
    os.remove(f)

URL = 'file:///home/geeker/architecture/home/geeker/code_repo/best_practice_perf/labs_ai_/agent_app/intro_self/index_believe.html'
PORT = 9339
PROF = '/tmp/believe_render_%d' % os.getpid()

chrome = subprocess.Popen([
    'google-chrome-stable',
    '--headless=new',
    '--enable-unsafe-swiftshader',
    '--use-gl=angle',
    '--use-angle=swiftshader',
    '--remote-debugging-port=%d' % PORT,
    '--user-data-dir=%s' % PROF,
    '--window-size=1600,900',
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-dev-shm-usage',
    'about:blank'
], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

try:
    # wait for DevTools to come up
    import urllib.request
    for _ in range(30):
        try:
            r = urllib.request.urlopen('http://localhost:%d/json' % PORT, timeout=1)
            data = json.loads(r.read())
            if data:
                break
        except Exception:
            time.sleep(0.4)
    else:
        print('DevTools never came up')
        raise SystemExit(1)

    # find the page target (not background_page)
    targets = json.loads(urllib.request.urlopen('http://localhost:%d/json' % PORT).read())
    page = None
    for t in targets:
        if t.get('type') == 'page':
            page = t
            break
    if not page:
        # the first page is about:blank; use it
        page = targets[0]

    ws_url = page['webSocketDebuggerUrl']
    ws = websocket.create_connection(ws_url, timeout=15)
    mid = [0]

    def send(method, params=None):
        mid[0] += 1
        msg = {'id': mid[0], 'method': method}
        if params:
            msg['params'] = params
        ws.send(json.dumps(msg))
        while True:
            resp = json.loads(ws.recv())
            if resp.get('id') == mid[0]:
                if 'error' in resp:
                    print('CDP error:', resp['error'])
                return resp

    # navigate to the deck
    send('Page.enable')
    send('Runtime.enable')
    send('Page.navigate', {'url': URL})
    time.sleep(2.5)

    # set device metrics
    send('Emulation.setDeviceMetricsOverride', {
        'width': 1600, 'height': 900, 'deviceScaleFactor': 1, 'mobile': False
    })
    time.sleep(1.0)

    # get total slides
    n_expr = 'document.querySelectorAll(".slide").length'
    nr = send('Runtime.evaluate', {'expression': n_expr, 'returnByValue': True})
    total = int(nr['result']['result']['value'])
    print('Total slides:', total)

    results = []
    for i in range(total):
        # navigate to slide i (goto(0-based))
        nav_expr = 'document.querySelectorAll(".slide").%d.classList.contains("active")' % i
        # use the global goto function
        send('Runtime.evaluate', {'expression': 'goto(%d)' % i, 'returnByValue': True})
        time.sleep(1.6)  # wait for reveal animations

        shot_path = os.path.join(SHOTS, 'slide_%02d.png' % (i+1))
        r = send('Page.captureScreenshot', {
            'format': 'png',
            'captureBeyondViewport': False,
            'fromSurface': True
        })
        if 'result' in r and 'data' in r['result']:
            import base64
            with open(shot_path, 'wb') as f:
                f.write(base64.b64decode(r['result']['data']))
            sz = os.path.getsize(shot_path)
            results.append((i+1, sz))
            print('slide %02d: %d bytes' % (i+1, sz))
        else:
            print('slide %02d: FAILED %s' % (i+1, r))

    ws.close()
    print('\n%d/%d slides captured' % (len(results), total))
except Exception as e:
    import traceback
    traceback.print_exc()
finally:
    chrome.terminate()
    try:
        chrome.wait(timeout=5)
    except Exception:
        chrome.kill()
    shutil.rmtree(PROF, ignore_errors=True)
