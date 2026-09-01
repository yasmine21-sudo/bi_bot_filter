import sys
import main
from main import RecipientGroup

def mock_load_recipient_mappings(path="destinataire.xlsx"):
    return {
        "gamma-security": RecipientGroup(
            to=["abdelkaderyasser.djemil@groupe-hasnaoui.com"],
            cc=[],
            original_name="GAMMA SECURITY"
        )
    }

main.load_recipient_mappings = mock_load_recipient_mappings

if __name__ == "__main__":
    try:
        sys.exit(main.main())
    except Exception as error:
        print(f"Test run failed with error: {error}")
        sys.exit(1)
