from inaSpeechSegmenter import Segmenter
import sys

# pip install inaSpeechSegmenter


if __name__ == '__main__':
    seg_model = Segmenter(vad_engine='smn', detect_gender=True)
    if len(sys.argv) != 3:
        print('usage .py path2wav.wav path2out_text.txt')
        exit()
    input_path = sys.argv[1].replace('\\', '/')
    output_path = sys.argv[2].replace('\\', '/')
    seg_data = seg_model(input_path)
    with open(output_path, 'w') as f:
        for d in seg_data:
            f.write('{},{},{}\n'.format(d[0], d[1], d[2]))


