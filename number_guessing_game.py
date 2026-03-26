import random


def get_hint(guess, answer):
    diff = abs(guess - answer)
    if diff == 0:
        return "正解!"
    elif diff <= 5:
        return "非常に近い!"
    elif diff <= 15:
        return "近い!"
    elif diff <= 30:
        return "まあまあ近い"
    else:
        direction = "大きい" if guess > answer else "小さい"
        return f"遠い (答えはもっと{'小さい' if guess > answer else '大きい'})"


def play():
    answer = random.randint(1, 100)
    max_attempts = 10
    attempts = 0

    print("=" * 40)
    print("  数字当てゲーム (1〜100)")
    print("=" * 40)
    print(f"1から100の数字を当ててください。チャンスは{max_attempts}回です。\n")

    while attempts < max_attempts:
        remaining = max_attempts - attempts
        print(f"残り{remaining}回", end=" | ")

        try:
            guess = int(input("数字を入力: "))
        except ValueError:
            print("  数字を入力してください。\n")
            continue

        if guess < 1 or guess > 100:
            print("  1〜100の範囲で入力してください。\n")
            continue

        attempts += 1
        hint = get_hint(guess, answer)

        if guess == answer:
            print(f"\n  {hint} 正解です! 答えは {answer} でした。")
            print(f"  {attempts}回で当てました!")
            break

        direction = "↓ もっと小さい" if guess > answer else "↑ もっと大きい"
        print(f"  {hint} → {direction}\n")

    else:
        print(f"\n  残念! 正解は {answer} でした。")

    print()
    play_again = input("もう一度プレイしますか? (y/n): ").strip().lower()
    if play_again == "y":
        print()
        play()
    else:
        print("ありがとうございました!")


if __name__ == "__main__":
    play()
