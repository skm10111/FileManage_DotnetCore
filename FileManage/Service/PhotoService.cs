using CloudinaryDotNet;
using CloudinaryDotNet.Actions;
using FileManage.Interface;
using Microsoft.Extensions.Options;

namespace FileManage.Service
{
    public class PhotoService : IPhotoService
    {
        private readonly Cloudinary _cloudinary;

        public class CloudinarySettings
        {
            public string CloudName { get; set; } 
            public string APIKey { get; set; }
            public string APISecret { get; set; }
        }
        public PhotoService(IOptions<CloudinarySettings> config)
        {
            var acc = new Account
           (
               config.Value.CloudName,
               config.Value.APIKey,
               config.Value.APISecret
           );
            _cloudinary = new Cloudinary(acc);
        }

        public async Task<ImageUploadResult> AddPhotoAsync(IFormFile file)
        {
            var uploadResult = new ImageUploadResult();

            if (file.Length > 0)
            {
                using var stream = file.OpenReadStream();
                var uploadParams = new ImageUploadParams
                {
                    File = new FileDescription(file.FileName, stream),
                    Transformation = new Transformation().Height(500).Width(500).Crop("fill").Gravity("face")
                };
                uploadResult = await _cloudinary.UploadAsync(uploadParams);
            }

            return uploadResult;
        }

        public async Task<DeletionResult> DeletePhotoAsync(string publicId)
        {

            var deleteParams = new DeletionParams(publicId);

            var result = await _cloudinary.DestroyAsync(deleteParams);

            return result;
        }
    }
}
