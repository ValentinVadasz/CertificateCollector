import win32com.client
import pprint
import os

thumbprints = {}
for subdir, dirs, files in os.walk(r'C:\Users\Administrator.ROLES1\Desktop\SecurityAudit\Application\7.3.3'):
    for file in files:
        filepath = subdir + os.sep + file

        print(filepath)
        s = win32com.client.gencache.EnsureDispatch('capicom.signedcode',0)
        s.FileName = filepath
        try:
            signer = s.Signer
            thumbprints[signer.Certificate.Thumbprint] = signer.Certificate
            print("Subject: ", signer.Certificate.SubjectName)
            print("Issuer: ", signer.Certificate.IssuerName)
            print("Thumbprint: ", signer.Certificate.Thumbprint)
            print("SerialNumber: ", signer.Certificate.SerialNumber + "\n\n")
        except:
            print("Nincs aláírva\n\n")

print("\n\n\n=================================================================================================================")
for key in thumbprints:
    print("Subject: ", thumbprints[key].SubjectName)
    print("Issuer: ", thumbprints[key].IssuerName)
    print("Thumbprint: ", thumbprints[key].Thumbprint)
    print("SerialNumber: ", thumbprints[key].SerialNumber + "\n\n")
print("=================================================================================================================")



