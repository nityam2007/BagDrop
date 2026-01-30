import os
from pathlib import Path
from groq import Groq

# Initialize Groq client with API key
client = Groq(api_key="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")

# Output directory for voice files
output_dir = Path(__file__).parent / "voice_files"
output_dir.mkdir(exist_ok=True)

# Scene voiceovers - CORRECT FLOW
# Scene 5: Store = FREE (locker opens, bag stored)
# Scene 6: Retrieve = Kiosk calculates time, pay, then collect
scenes = {
    "scene_01_arrival": "[confident] Arjun arrives at the railway station with his luggage.",
    
    "scene_02_wait": "[concerned] His flight is in five hours. [thoughtful] He cannot carry bags the entire time.",
    
    "scene_03_struggle": "[strained] Dragging luggage everywhere is tiring. [serious] He needs a better option.",
    
    "scene_04_discovery": "[cheerful] He finds BagDrop on his phone. [confident] A smart locker is nearby.",
    
    "scene_05_kiosk": "[calm] Arjun walks to the BagDrop kiosk. [professional] He selects an available locker and stores his bag. [reassuring] The locker closes securely.",
    
    "scene_06_retrieval": "[calm] When Arjun returns, the kiosk calculates the storage time. [informative] He sees the amount on screen, pays using his card, and the locker opens. [cheerful] He collects his bag.",
    
    "scene_07_status": "[informative] Green means available. Red means occupied. [confident] Simple and clear.",
    
    "scene_08_closing": "[confident] BagDrop. [cheerful] Store your luggage. Travel freely."
}

print("=" * 50)
print("BagDrop Voice Generation Script")
print("=" * 50)
print(f"Model: canopylabs/orpheus-v1-english")
print(f"Voice: hannah")
print(f"Output: {output_dir}")
print("=" * 50)

# Generate audio for each scene
for scene_name, text in scenes.items():
    output_file = output_dir / f"{scene_name}.wav"
    print(f"\nGenerating: {scene_name}.wav")
    print(f"  Text: {text[:60]}...")
    
    try:
        response = client.audio.speech.create(
            model="canopylabs/orpheus-v1-english",
            voice="hannah",
            response_format="wav",
            input=text,
        )
        
        # Use the correct method from Groq SDK
        response.write_to_file(str(output_file))
        
        print(f"  ✓ Saved: {output_file}")
    except Exception as e:
        print(f"  ✗ Error: {e}")

print("\n" + "=" * 50)
print("Done! All voice files generated.")
print(f"Location: {output_dir}")
print("=" * 50)
