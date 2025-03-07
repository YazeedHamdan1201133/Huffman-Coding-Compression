import heapq
import math
import docx2txt

# Read text from a .docx file
def read_text_from_docx(file_name):
    text = docx2txt.process(file_name)
    return text.lower().replace("\n", "")  # Convert to lowercase and remove newlines

# Build a frequency dictionary
def calculate_frequency(text):
    frequency = {}
    for char in text:
        if char == ' ':
            char = '(space)'
        if char not in frequency:
            frequency[char] = 0
        frequency[char] += 1
    return frequency

# Calculate probabilities
def calculate_probabilities(frequency, total_characters):
    probabilities = {char: freq / total_characters for char, freq in frequency.items()}
    return probabilities

# Calculate entropy
def calculate_entropy(probabilities):
    entropy = -sum(prob * math.log2(prob) for prob in probabilities.values())
    return entropy

# Build Huffman tree
class Node:
    def __init__(self, char=None, frequency=0, left=None, right=None):
        self.char = char
        self.frequency = frequency
        self.left = left
        self.right = right

    def __lt__(self, other):
        return self.frequency < other.frequency

def build_huffman_tree(frequency):
    heap = [Node(char, freq) for char, freq in frequency.items()]
    heapq.heapify(heap)

    while len(heap) > 1:
        left = heapq.heappop(heap)
        right = heapq.heappop(heap)
        merged = Node(frequency=left.frequency + right.frequency, left=left, right=right)
        heapq.heappush(heap, merged)

    return heap[0]

# Generate Huffman codes
def generate_huffman_codes(node, code="", codes=None):
    if codes is None:
        codes = {}
    if node.char is not None:
        codes[node.char] = code
    if node.left:
        generate_huffman_codes(node.left, code + "0", codes)
    if node.right:
        generate_huffman_codes(node.right, code + "1", codes)
    return codes

# Calculate average bits per character
def calculate_average_bits(probabilities, codes):
    average_bits = sum(probabilities[char] * len(code) for char, code in codes.items())
    return average_bits

# Calculate total bits for Huffman and ASCII
def calculate_total_bits(frequency, codes):
    nhuffman = sum(frequency[char] * len(code) for char, code in codes.items())
    nascii = sum(frequency.values()) * 8  # ASCII uses 8 bits per character
    return nascii, nhuffman

# Main function
if __name__ == "__main__":
    file_name = r"The_Path"
    text = read_text_from_docx(file_name)

    # Step 1: Calculate frequency and total characters
    frequency = calculate_frequency(text)
    total_characters = sum(frequency.values())

    # Step 2: Calculate probabilities and entropy
    probabilities = calculate_probabilities(frequency, total_characters)
    entropy = calculate_entropy(probabilities)

    # Step 3: Build Huffman tree and generate codes
    huffman_tree = build_huffman_tree(frequency)
    codes = generate_huffman_codes(huffman_tree)

    # Step 4: Calculate average bits and total bits
    average_bits = calculate_average_bits(probabilities, codes)
    nascii, nhuffman = calculate_total_bits(frequency, codes)

    # Step 5: Calculate compression percentage
    compression_percentage = (nascii - nhuffman) / nascii * 100

    # Display results
    print("\n\n-------------------------")
    print("Total Characters:", total_characters)
    print("-------------------------")
    # Display a unified table of results
    print("\nCharacter | Frequency | Probability | Codeword       | Length")
    print("------------------------------------------------------------------")
    for char, code in sorted(codes.items()):
        char_display = char if char != "(space)" else "space"
        print(f"{char_display:<10} | {frequency[char]:<10} | {probabilities[char]:.5f}     | {code:<12} | {len(code):<6}")
    print("------------------------------------------------------------------")
    print("Total Bits (NASCII): {:.0f} bits".format(nascii))
    print("Average Bits (Huffman): {:.5f} bits/character".format(average_bits))
    print("Entropy: {:.5f} bits/character".format(entropy))
    print("Total Bits (Huffman): {:.0f} bits".format(nhuffman))
    print("Compression Percentage: {:.2f}%".format(compression_percentage))
    # Table for specific symbols
    print("\nTable for Selected Symbols:")
    print("---------------------------------------------------------------")
    # Adjust the column widths
    print("{:<10} {:<16} {:<18} {:>10}".format("Symbol", "Probability", "Codeword", "Length"))
    print("---------------------------------------------------------------")
    for char in ['a', 'b', 'c', 'd', 'e', 'f', 'm', 'z', '(space)', '.']:
        if char in codes:
            print(f"{char:<10} {probabilities[char]:<16.5f} {codes[char]:<18} {len(codes[char]):>10}")
    print("---------------------------------------------------------------\n\n")

