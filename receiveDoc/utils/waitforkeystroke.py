import time
from pynput import keyboard
from threading import Timer

# 计时器调用函数
# def hello(*args):
#     starttime=args[0]
#     curtime = time.perf_counter()
#     elapsed = curtime - starttime
#     print(f"time elapsed: {elapsed}")
#     s=args[1]
#     print(s)


def secondcounter(*args):
    print("timer in effect")
    seconds=args[0]
    for i in range(seconds):
        time.sleep(1)
        print(f"\t \t another {i+1} seconds passed...")
    keyboard.Controller().press(keyboard.Key.enter)
    keyboard.Controller().release(keyboard.Key.enter)


# # 键盘监听
# def on_press(key):
#     if key == keyboard.Key.enter:
#         print("timer cancelled")
#         timer_keystroke.cancel()


# Warning: 在on_release中返回False才有效，在on_press中返回False无法终止监听
# def on_release(key):
#     if key == keyboard.Key.enter:
#         print("listener stopped")
#         return False


# 启动计时器
# starttime = time.perf_counter()
# timer_keystroke = Timer(interval=10, function=hello, args=[starttime,"func within timer"])
# interval=3
# duration=5


def waitforkeystroke(interval=3, duration=3):
    timer_keystroke = Timer(interval=interval, function=secondcounter, args=(duration,))
    timer_keystroke.start()
    print(f"timer started, interval {interval}s")
    # print("to cancel the timer, press Enter")
    # print("to terminate keyboard listener, press Enter")
    # 键盘监听  # Collect events until released
    def on_press(key):
        if key == keyboard.Key.enter:
            print("timer cancelled")
            timer_keystroke.cancel()

    def on_release(key):
        if key == keyboard.Key.enter:
            print("listener stopped")
            return False

    with keyboard.Listener(
            on_press=on_press,
            on_release=on_release) as listener:
        listener.join()




if __name__ == "__main__":
    # timer_keystroke = Timer(interval=3, function=secondcounter, args=(5,))
    waitforkeystroke()
    print("move on to next line")
    print("type to confirm and submit...")
