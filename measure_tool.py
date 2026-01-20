import cv2
import numpy as np
import math
import os
import sys

# グローバル変数
points = []
img_display = None
img_clean = None
scale_ratio = 1.0
is_calibrated = False

def click_event(event, x, y, flags, param):
    global points, img_display

    if event == cv2.EVENT_LBUTTONDOWN:
        points.append((x, y))
        cv2.circle(img_display, (x, y), 3, (0, 0, 255), -1)
        if len(points) > 1:
            cv2.line(img_display, points[-2], points[-1], (0, 0, 255), 1)
        cv2.imshow('Plant Analysis Tool', img_display)

def main(image_path, known_scale_um):
    global img_display, img_clean, points, scale_ratio, is_calibrated

    # 画像読み込みチェック
    if not os.path.exists(image_path):
        print(f"エラー: ファイル '{image_path}' が見つかりません。")
        print("正しいパスを指定してください。")
        return

    img = cv2.imread(image_path)
    img_clean = img.copy()
    img_display = img.copy()

    print("=== 植物画像解析ツール ===")
    print(f"解析対象: {image_path}")
    print(f"スケール設定値: {known_scale_um} μm")
    print("------------------------------------------------")
    print("【ステップ1: スケール調整】")
    print(f"画像のスケールバー（{known_scale_um}μm）の左端と右端をクリックしてください。")

    cv2.namedWindow('Plant Analysis Tool', cv2.WINDOW_NORMAL)
    cv2.setMouseCallback('Plant Analysis Tool', click_event)
    cv2.imshow('Plant Analysis Tool', img_display)

    while True:
        key = cv2.waitKey(1) & 0xFF

        # --- キャリブレーション ---
        if not is_calibrated and len(points) == 2:
            p1, p2 = points[0], points[1]
            px_dist = math.sqrt((p1[0]-p2[0])**2 + (p1[1]-p2[1])**2)

            if px_dist > 0:
                scale_ratio = known_scale_um / px_dist # ここで指定された値を使用
                is_calibrated = True

                # 完了表示
                cv2.line(img_display, p1, p2, (0, 255, 255), 2)
                cv2.putText(img_display, f"Scale: {known_scale_um}um", (p1[0], p1[1]-10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,255,255), 2)
                cv2.imshow('Plant Analysis Tool', img_display)

                print(f"完了: 1 pixel = {scale_ratio:.4f} μm")
                print("\n【ステップ2: 測定モード】")
                print("  [d]キー : 距離測定 (2点クリック → d)")
                print("  [a]キー : 面積測定 (囲んで → a)")
                print("  [r]キー : リセット")
                print("  [q]キー : 終了")
                points = []
            else:
                points = []

        # --- [d] 距離測定 ---
        elif key == ord('d'):
            if len(points) >= 2:
                p_start, p_end = points[-2], points[-1]
                dist_px = math.sqrt((p_start[0]-p_end[0])**2 + (p_start[1]-p_end[1])**2)
                dist_um = dist_px * scale_ratio

                cv2.line(img_display, p_start, p_end, (0, 255, 0), 2)
                mid_x, mid_y = (p_start[0] + p_end[0]) // 2, (p_start[1] + p_end[1]) // 2
                cv2.putText(img_display, f"{dist_um:.1f} um", (mid_x, mid_y),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
                print(f"距離: {dist_um:.2f} μm")
                points = []
                cv2.imshow('Plant Analysis Tool', img_display)

        # --- [a] 面積測定 ---
        elif key == ord('a'):
            if len(points) >= 3:
                pts_array = np.array(points, np.int32).reshape((-1, 1, 2))
                area_px = cv2.contourArea(pts_array)
                area_um = area_px * (scale_ratio ** 2)

                overlay = img_display.copy()
                cv2.fillPoly(overlay, [pts_array], (255, 0, 255))
                cv2.addWeighted(overlay, 0.4, img_display, 0.6, 0, img_display)
                cv2.polylines(img_display, [pts_array], True, (255, 0, 255), 2)

                M = cv2.moments(pts_array)
                if M['m00'] != 0:
                    cx, cy = int(M['m10']/M['m00']), int(M['m01']/M['m00'])
                    cv2.putText(img_display, f"{area_um:.0f} um2", (cx-40, cy),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)

                print(f"面積: {area_um:.2f} μm²")
                points = []
                cv2.imshow('Plant Analysis Tool', img_display)

        elif key == ord('r'):
            img_display = img_clean.copy()
            points = []
            print("リセットしました")
            cv2.imshow('Plant Analysis Tool', img_display)

        elif key == ord('q'):
            break

    cv2.destroyAllWindows()

if __name__ == "__main__":
    # --- コマンドライン引数の処理 ---

    # 引数が足りない場合
    if len(sys.argv) < 2:
        print("使い方: python measure_tool.py <画像ファイル名> [スケール値(um)]")
        print("例1: python measure_tool.py C4-1023-3.jpg")
        print("例2: python measure_tool.py C4-1023-3.jpg 100")
        sys.exit() # 終了

    # 第1引数: ファイル名
    target_file = sys.argv[1]

    # 第2引数: スケール値 (なければ50.0)
    target_scale = 50.0
    if len(sys.argv) > 2:
        try:
            target_scale = float(sys.argv[2])
        except ValueError:
            print("エラー: スケール値は数値を入力してください。")
            sys.exit()

    # メイン処理実行
    main(target_file, target_scale)
