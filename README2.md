look at the .osts file. how can a python script be created to do the similar task?

You can replicate that .osts script in Python by using an Excel library (like `openpyxl` for files on disk or `pywin32` for driving a live Excel instance). Below is a concrete `openpyxl` example plus a `pywin32` variant.

## What the .osts script does

From the `Script-2.osts` content, the Office Script does this on the active worksheet: [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)

- Gets the active worksheet. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)
- Inserts a column at `MB:MB`, shifting existing columns to the right. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)
- Hides column `MB`. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)
- Sets cell `MC1` to the text `"COMMENTS 10-03-2026"`. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)

(Ignoring the `copyFrom` call here, because in the recorded script it effectively copies the range to itself and does not add extra behavior in a typical replay context. ) [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)

## Python script using openpyxl (workbook file on disk)

Use this if you want to modify a `.xlsx` file directly, not a live Excel window.

```python
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

def update_comments_column(
    file_path: str,
    sheet_name: str | None = None,
):
    # Load workbook
    wb = load_workbook(file_path)

    # Pick sheet: given name or active sheet
    if sheet_name is not None:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    # 1) Insert a column at MB (shift right)
    mb_index = column_index_from_string("MB")
    ws.insert_cols(mb_index)

    # 2) Hide column MB
    ws.column_dimensions["MB"].hidden = True

    # 3) Set MC1 header text with today's date in the same format
    today_str = datetime.today().strftime("%d-%m-%Y")
    ws["MC1"] = f"COMMENTS {today_str}"

    # Save changes (overwrite or new file)
    wb.save(file_path)

if __name__ == "__main__":
    # Example usage
    update_comments_column(r"C:\path\to\your\workbook.xlsx", sheet_name=None)
```

This script mirrors the logic of the Office Script but operates on a given workbook file instead of the Excel "active worksheet" concept. [ppl-ai-file-upload.s3.amazonaws](https://ppl-ai-file-upload.s3.amazonaws.com/web/direct-files/attachments/19012317/a3f79f33-bb69-47dd-a65f-905566cf78dc/Script-2.osts?AWSAccessKeyId=ASIA2F3EMEYE7KOE7VMF&Signature=esxnTgmPLT5dDpUjEZdqns2Nfn8%3D&x-amz-security-token=IQoJb3JpZ2luX2VjEH4aCXVzLWVhc3QtMSJIMEYCIQC%2BAnNWxtnXEnsjm0ust9HxTo6Xn8AIPdWfkIJJtCMF4QIhAL1Kjhg3tr%2BRElzF0NtwZ%2FLOf6Nlb0VxNJHZFEwblyMvKvMECEcQARoMNjk5NzUzMzA5NzA1Igx%2BI6SbVPJndAgVCLYq0ATrdmQRjb4dnJA2GtIPkOfPwLjtYRbDNckmlv6ZdmRWsRJY0XQoe1FNXauMqQD0e4g6lfW4j9W4OUG0r832%2FrYUPdiCNsJJ6WIipZ1p2LpO9MSfuU8txmNoIChyxo%2BQR0zYgfAi62eXwu31iCPMmGsMR%2BkkdPBNZlvs9UyobCBMzeQXwVBRqEHuDA5X9DKo8Ne4R5%2BjklT%2FWaTn15JOHznsSs5eU0fTsTpZYxZTtLS8Ed2r1hN762egrte6FjuPvw4SotWxc2qdsKmPOB0hmyQAb67t6yopHeiJOYPmAsHoUqdWkj4sDFT9V%2Fz7y0LuGuqIvdWpBZxXipHpYkgDjHoGdpWyP48TAc0jL534oNSvA5SQdeGENZErN%2F9aAtMQ0gzY%2BsmB%2FR1CzgIHAq16mAm7UHcXcGGHfamADUor8dsf1Sos%2FZ%2BhNAS2mT4Q8vxBfv%2BHEnP52fR6PibbSU74q3%2FJykTwXIIdrWMjv79ewJ4Yk31nUKuAWZ%2FTtVbGD7U%2FeIkNY91HaQmm2r43ykBZ02tISYCrYJgklwlq88hj60EqVlAQIDQOTandvYv7Yf94KjuiKYIqdohZd%2BFWW4ZCE87BkQ2VevY7bm%2F5zYolFeXlwyL%2FqIlV79FBAzgIp3S5oN8SGuw6ijcZxih8Fm9vtmTEiH12QT0dabdUncu%2BxPmmfXXUXvumIwjHVE6bB2yC8PxirssrfO%2FnUIQRKwwCEXP4nYNsDBmv1hshpxewWi36v9fJw%2Frss95zlaDX5TNEabWmTnBFMkmCYueytSPI0AcDMJu4wM0GOpcButVtYxUAl6QAXJL%2B0AWZh6lPMObHGM2iAgyurErLJoWYtOYCoO6bnjuuM84sLZD86FkPwiZYbamvMG%2BWg1KuRJDEnjAGHMeR%2FWF0aW547y%2Bb7%2BaGr4%2FdYA46bbphzDNIAAsCmKtDAJW96lZSYg68CqMkBJ4KgovdmgkTPsWhIc7TgEOqxoD8mxUvfAsfdxXqT6oiasnADw%3D%3D&Expires=1773151256)

## Python script using pywin32 (control live Excel)

If you want to drive the currently-open workbook and its active sheet (closer to what the Office Script does in Excel Online/desktop), you can use `pywin32`:

```python
import pythoncom
from datetime import datetime
from win32com.client import Dispatch, constants

def update_comments_in_active_excel():
    pythoncom.CoInitialize()  # Ensure COM is initialized

    excel = Dispatch("Excel.Application")
    wb = excel.ActiveWorkbook
    ws = excel.ActiveSheet  # active worksheet

    # 1) Insert column MB, shift right
    ws.Range("MB:MB").EntireColumn.Insert(Shift=constants.xlToRight)

    # 2) Hide column MB
    ws.Columns("MB:MB").Hidden = True

    # 3) Set MC1 value with today's date
    today_str = datetime.today().strftime("%d-%m-%Y")
    ws.Range("MC1").Value = f"COMMENTS {today_str}"

    # Optionally save
    wb.Save()

if __name__ == "__main__":
    update_comments_in_active_excel()
```

This assumes:

- Excel is open.
- The workbook and sheet you want are currently active.
- You have `pywin32` installed (`pip install pywin32`).

Would you prefer modifying saved `.xlsx` files on disk, or directly automating a workbook that is already open in Excel on your machine?

Perfect! Here's a production-ready Python script for modifying saved `.xlsx` files on disk using `openpyxl`:

```python
"""
Insert and hide a column, then add a dated comment header.
Replicates the behavior of Script-2.osts for local Excel files.
"""

import argparse
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def insert_comments_column(
    file_path: str | Path,
    sheet_name: str | None = None,
    insert_column: str = "MB",
    header_column: str = "MC",
    date_format: str = "%d-%m-%Y",
    backup: bool = False,
) -> None:
    """
    Insert a new column, hide it, and add a dated header.

    Args:
        file_path: Path to the .xlsx file
        sheet_name: Target sheet name (None = active sheet)
        insert_column: Column letter to insert at (default: MB)
        header_column: Column letter for the header cell (default: MC)
        date_format: Date format string for header (default: %d-%m-%Y)
        backup: Create a backup copy before modifying
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Optional backup
    if backup:
        backup_path = file_path.with_suffix(f".backup{file_path.suffix}")
        import shutil
        shutil.copy2(file_path, backup_path)
        print(f"✓ Backup created: {backup_path}")

    # Load workbook
    wb = load_workbook(file_path)
    
    # Select sheet
    ws = wb[sheet_name] if sheet_name else wb.active
    print(f"✓ Working on sheet: {ws.title}")

    # 1. Insert column at specified position (shifts right)
    col_index = column_index_from_string(insert_column)
    ws.insert_cols(col_index)
    print(f"✓ Inserted column at {insert_column}")

    # 2. Hide the inserted column
    ws.column_dimensions[insert_column].hidden = True
    print(f"✓ Hidden column {insert_column}")

    # 3. Set header with today's date
    today_str = datetime.today().strftime(date_format)
    header_cell = f"{header_column}1"
    ws[header_cell] = f"COMMENTS {today_str}"
    print(f"✓ Set {header_cell} = 'COMMENTS {today_str}'")

    # Save changes
    wb.save(file_path)
    print(f"✓ Saved: {file_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Insert and hide a column with dated comment header"
    )
    parser.add_argument(
        "file",
        type=str,
        help="Path to the Excel .xlsx file"
    )
    parser.add_argument(
        "-s", "--sheet",
        type=str,
        default=None,
        help="Sheet name (default: active sheet)"
    )
    parser.add_argument(
        "-i", "--insert-column",
        type=str,
        default="MB",
        help="Column to insert at (default: MB)"
    )
    parser.add_argument(
        "-c", "--header-column",
        type=str,
        default="MC",
        help="Column for header cell (default: MC)"
    )
    parser.add_argument(
        "-b", "--backup",
        action="store_true",
        help="Create backup before modifying"
    )
    parser.add_argument(
        "-d", "--date-format",
        type=str,
        default="%d-%m-%Y",
        help="Date format for header (default: %%d-%%m-%%Y)"
    )

    args = parser.parse_args()

    try:
        insert_comments_column(
            file_path=args.file,
            sheet_name=args.sheet,
            insert_column=args.insert_column,
            header_column=args.header_column,
            date_format=args.date_format,
            backup=args.backup,
        )
        print("\n✅ Complete!")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        raise


if __name__ == "__main__":
    main()
```

## Installation

```bash
pip install openpyxl
```

## Usage Examples

**Basic usage (modifies file in-place):**
```bash
python insert_comments_column.py "C:\data\workbook.xlsx"
```

**With backup:**
```bash
python insert_comments_column.py "C:\data\workbook.xlsx" --backup
```

**Target specific sheet:**
```bash
python insert_comments_column.py "C:\data\workbook.xlsx" --sheet "Sheet2"
```

**Custom columns:**
```bash
python insert_comments_column.py "C:\data\workbook.xlsx" --insert-column "AA" --header-column "AB"
```

**Different date format:**
```bash
python insert_comments_column.py "C:\data\workbook.xlsx" --date-format "%Y-%m-%d"
```

## Key Features

✅ **Exact .osts replication** - Inserts column at MB, hides it, sets MC1 header  
✅ **CLI with argparse** - Flexible command-line interface  
✅ **Optional backup** - Protects original file  
✅ **Configurable columns** - Not hardcoded to MB/MC  
✅ **Date formatting** - Customizable date format  
✅ **Error handling** - File validation and informative messages  
✅ **Cross-platform** - Works on Windows/Linux/Mac  

This gives you the same functionality as your Office Script but for local `.xlsx` files on disk!